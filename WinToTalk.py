# -*- coding: utf-8 -*-

import asyncio
import json
import sys
import signal
import threading
import queue
from datetime import datetime
from dataclasses import dataclass

import websockets
import comtypes.client
import pythoncom


SVS_ASYNC = 1
SVS_PURGE = 2

DEFAULT_RATE = 300
DEFAULT_VOLUME = 100


@dataclass
class SpeechItem:
    text: str
    language: str
    gender: str
    rate: int
    volume: int
    speaker: str


speech_queue = queue.Queue()

cancel_event = threading.Event()
stop_event = threading.Event()


def rate_to_sapi(rate):
    baseline = 200
    diff = rate - baseline
    return max(-10, min(10, int(diff / 20)))


def select_voice(voice, language, gender):

    name = "Microsoft Susan"

    if gender.lower() == "male":
        name = "Microsoft Richard"

    for v in voice.GetVoices():
        if name.lower() in v.GetDescription().lower():
            voice.Voice = v
            return


def tts_worker():

    pythoncom.CoInitialize()

    voice = comtypes.client.CreateObject("SAPI.SpVoice")

    print("[TTS] Worker started")

    try:

        while not stop_event.is_set():

            print("[TTS] Waiting for queue item...")

            item = speech_queue.get()

            if item is None:
                break

            print(f"[TTS] QUEUE GET ({item.speaker}) | size={speech_queue.qsize()}")

            voice.Volume = item.volume
            voice.Rate = rate_to_sapi(item.rate)

            select_voice(voice, item.language, item.gender)

            print(f"[TTS] SPEAK START ({item.speaker})")

            cancel_event.clear()

            voice.Speak(item.text, SVS_ASYNC)

            while True:

                # wait 100 ms for completion
                finished = voice.WaitUntilDone(100)

                if finished:
                    print("[TTS] SPEAK FINISHED")
                    break

                if cancel_event.is_set():

                    print("[TTS] CANCEL EXECUTED")

                    voice.Speak("", SVS_PURGE)

                    cancel_event.clear()
                    break

            speech_queue.task_done()

    finally:

        pythoncom.CoUninitialize()


worker_thread = threading.Thread(target=tts_worker, daemon=True)
worker_thread.start()


def enqueue_speech(text, language, gender, rate, volume, speaker):

    item = SpeechItem(text, language, gender, rate, volume, speaker)

    speech_queue.put(item)

    print(f"[TTS] QUEUE PUT ({speaker}) | size={speech_queue.qsize()}")


def cancel_current():

    print("[TTS] CANCEL REQUESTED")

    cancel_event.set()


async def process_message(msg):

    try:

        data = json.loads(msg)

        msg_type = data.get("Type", "").lower()

        if msg_type == "say":

            payload = data.get("Payload", "")
            language = data.get("Language", "English")

            gender = data.get("Voice", {}).get("Name", "None")

            rate = data.get("Rate", DEFAULT_RATE)
            speaker = data.get("Speaker", "Unknown")

            print(datetime.now(), "Say", speaker)

            enqueue_speech(payload, language, gender, rate, DEFAULT_VOLUME, speaker)

        elif msg_type == "cancel":

            cancel_current()

    except Exception as e:

        print("Message error:", e)


async def websocket_loop(uri):

    while True:

        try:

            async with websockets.connect(uri) as ws:

                print("[WinToTalk] Connected")

                while True:

                    msg = await ws.recv()

                    await process_message(msg)

        except Exception as e:

            print("WebSocket error:", e)

            await asyncio.sleep(1)


def shutdown(*args):

    print("[WinToTalk] Shutdown")

    stop_event.set()
    cancel_event.set()

    speech_queue.put(None)

    worker_thread.join(timeout=2)

    sys.exit(0)


if __name__ == "__main__":

    URI = "ws://localhost:3000/Messages"

    signal.signal(signal.SIGINT, shutdown)
    signal.signal(signal.SIGTERM, shutdown)

    asyncio.run(websocket_loop(URI))
