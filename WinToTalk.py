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
import re
from langdetect import detect, DetectorFactory

DEFAULT_RATE = 300
DEFAULT_VOLUME = 100

# ------------------------
# Voice configuration
# ------------------------
EN_NEUTRAL = "Microsoft Catherine"
EN_FEMALE  = "Microsoft Susan"
EN_MALE    = "Microsoft Richard"

DE_NEUTRAL = "Microsoft Hedda Desktop"
DE_FEMALE  = "Microsoft Katja"
DE_MALE    = "Microsoft Karsten"

SVS_ASYNC = 1
SVS_PURGE = 2

DetectorFactory.seed = 0

def detect_chat_language(text, default_language):

    text = text.strip()

    # sehr zuverlässiger Deutsch-Indikator
    if any(c in text for c in "äöüÄÖÜß"):
        print(f"[WinToTalk] (shortcut) Detect Language = German")
        return "German"

    try:
        lang = detect(text)

        if lang == "de":
            print(f"[WinToTalk] (detect) Detect Language = German")
            return "German"

        if lang == "en":
            print(f"[WinToTalk] (detect) Detect Language = English")
            return "English"

    except:
        pass

    print(f"[WinToTalk] (default) Detect Language = ", default_language)
    return default_language

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

def sanitize_for_sapi(text: str) -> str:
    # remove XML brackets (SAPI interprets as SSML)
    text = text.replace("<", "").replace(">", "")

    # remove control characters
    text = re.sub(r'[\x00-\x1F\x7F]', '', text)

    return text

def rate_to_sapi(rate):
    baseline = 200
    diff = rate - baseline
    return max(-10, min(10, int(diff / 20)))


# Mapping für Logging
VOICE_MAP = {
    "EN_NEUTRAL": EN_NEUTRAL,
    "EN_FEMALE": EN_FEMALE,
    "EN_MALE": EN_MALE,
    "DE_NEUTRAL": DE_NEUTRAL,
    "DE_FEMALE": DE_FEMALE,
    "DE_MALE": DE_MALE,
}

def select_voice(voice, language, gender):
    """Select one of the 6 preconfigured voices and log clearly."""
    if language.lower().startswith("german"):
        if gender.lower() == "male":
            chosen = "DE_MALE"
        elif gender.lower() == "female":
            chosen = "DE_FEMALE"
        else:
            chosen = "DE_NEUTRAL"
    else:
        if gender.lower() == "male":
            chosen = "EN_MALE"
        elif gender.lower() == "female":
            chosen = "EN_FEMALE"
        else:
            chosen = "EN_NEUTRAL"

    voice_name = VOICE_MAP[chosen]

    for v in voice.GetVoices():
        if voice_name.lower() in v.GetDescription().lower():
            voice.Voice = v
            print(f"[WinToTalk] Selected voice constant: {chosen} -> {voice_name}")
            return chosen, voice_name

    print(f"[WinToTalk] Warning: voice '{voice_name}' not found, using default")
    return chosen, "Default system voice"


def tts_worker():

    pythoncom.CoInitialize()

    voice = comtypes.client.CreateObject("SAPI.SpVoice")

    #print("[TTS] Worker started")

    try:

        while not stop_event.is_set():

            #print("[TTS] Waiting for queue item...")

            item = speech_queue.get()

            if item is None:
                break

            #print(f"[TTS] QUEUE GET ({item.speaker}) | size={speech_queue.qsize()}")

            voice.Volume = item.volume
            voice.Rate = rate_to_sapi(item.rate)

            language = detect_chat_language(item.text, item.language)
            select_voice(voice, language, item.gender)

            #print(f"[TTS] SPEAK START ({item.speaker})")

            cancel_event.clear()

            try:
                safe_text = sanitize_for_sapi(item.text)
                voice.Speak(safe_text, SVS_ASYNC)
            except Exception as e:
                print("[TTS] Recovering from SAPI error:", e)
                voice = comtypes.client.CreateObject("SAPI.SpVoice")
                continue

            while True:

                # wait 100 ms for completion
                finished = voice.WaitUntilDone(100)

                if finished:
                    #print("[TTS] SPEAK FINISHED")
                    break

                if cancel_event.is_set():

                    #print("[TTS] CANCEL EXECUTED")

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

    #print(f"[TTS] QUEUE PUT ({speaker}) | size={speech_queue.qsize()}")
    
    if speech_queue.qsize() > 100:
        print("[TTS] queue overflow, clearing")
        with speech_queue.mutex:
            speech_queue.queue.clear()


def cancel_current():

    #print("[TTS] CANCEL REQUESTED")

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

            print(datetime.now(), "Say")
            print("Language:", language)
            print("Gender:", gender)
            print("Speaker:", speaker)
            print("Rate:", rate)
            print("Text:", payload)
            print("")
            

            enqueue_speech(payload, language, gender, rate, DEFAULT_VOLUME, speaker)

        elif msg_type == "cancel":
            
            print(datetime.now(), "Cancel")

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
