# -*- coding: utf-8 -*-
"""
WinToTalk: Windows TTS WebSocket Listener
- Direct SAPI5 via comtypes
- 6 built-in voices (EN/DE, neutral/female/male)
- Per-message rate support
- Cancel stops immediately and new speech plays
"""

import asyncio
import json
import sys
import signal
from datetime import datetime
import threading

import comtypes.client
import websockets

# ------------------------
# Voice configuration
# ------------------------
EN_NEUTRAL = "Microsoft Catherine"
EN_FEMALE  = "Microsoft Susan"
EN_MALE    = "Microsoft Richard"

DE_NEUTRAL = "Microsoft Hedda Desktop"
DE_FEMALE  = "Microsoft Katja"
DE_MALE    = "Microsoft Karsten"

DEFAULT_RATE = 300  # default words per minute
DEFAULT_VOLUME = 100  # default volume

# ------------------------
# TTS Engine (SAPI5)
# ------------------------
SpVoice = comtypes.client.CreateObject("SAPI.SpVoice")
voice_lock = threading.Lock()
current_speak_thread = None

def rate_to_sapi(rate_wpm: int) -> int:
    baseline = 200  # SAPI standard
    diff = rate_wpm - baseline
    # lineare Umrechnung, +/-10 maximal
    sapi_rate = int(diff / 20)  # 20 WPM = 1 SAPI rate step
    return max(-10, min(10, sapi_rate))


# Mapping für Logging
VOICE_MAP = {
    "EN_NEUTRAL": EN_NEUTRAL,
    "EN_FEMALE": EN_FEMALE,
    "EN_MALE": EN_MALE,
    "DE_NEUTRAL": DE_NEUTRAL,
    "DE_FEMALE": DE_FEMALE,
    "DE_MALE": DE_MALE,
}

def select_voice(language, gender):
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

    for v in SpVoice.GetVoices():
        if voice_name.lower() in v.GetDescription().lower():
            SpVoice.Voice = v
            print(f"[WinToTalk] Selected voice constant: {chosen} -> {voice_name}")
            return chosen, voice_name

    print(f"[WinToTalk] Warning: voice '{voice_name}' not found, using default")
    return chosen, "Default system voice"


def speak(text, language="English", gender="None", rate=DEFAULT_RATE, volume=DEFAULT_VOLUME):
    """Speak text in a separate thread. Stops previous speech immediately."""
    global current_speak_thread

    def run():
        with voice_lock:
            SpVoice.Volume = max(0, min(100, volume))
            SpVoice.Rate = rate_to_sapi(rate)
            select_voice(language, gender)
            SpVoice.Speak(text, 1)  # 1 = SVSFlagsAsync

    # stop previous speech
    cancel_speech()

    current_speak_thread = threading.Thread(target=run, daemon=True)
    current_speak_thread.start()


def cancel_speech():
    """Immediately stop current speech."""
    with voice_lock:
        SpVoice.Speak("", 3)  # 3 = SVSFPurgeBeforeSpeak


# ------------------------
# WebSocket Handling
# ------------------------
async def process_message(msg):
    try:
        data = json.loads(msg)
        msg_type = data.get("Type", "").lower()

        if msg_type == "say":
            payload = data.get("Payload", "")
            language = data.get("Language", "English")
            voice_info = data.get("Voice", {})
            gender = voice_info.get("Name", "None")
            rate = data.get("Rate", DEFAULT_RATE)
            speaker = data.get("Speaker", "Unknown")

            print(datetime.now(), "Say")
            print("Language:", language)
            print("Gender:", gender)
            print("Speaker:", speaker)
            print("Rate:", rate)
            print("Text:", payload)
            print("")

            speak(payload, language, gender, rate)

        elif msg_type == "cancel":
            print(datetime.now(), "Cancel")
            cancel_speech()

    except Exception as e:
        print(f"[WinToTalk] Message error: {e}")


async def websocket_loop(uri):
    while True:
        try:
            async with websockets.connect(uri, ping_interval=10, ping_timeout=5) as ws:
                print("[WinToTalk] Connected to WebSocket")
                while True:
                    msg = await ws.recv()
                    await process_message(msg)
        except Exception as e:
            print(f"[WinToTalk] WebSocket error: {e}")
            await asyncio.sleep(1)


# ------------------------
# Main
# ------------------------
def shutdown(*args):
    print("[WinToTalk] Shutting down...")
    cancel_speech()
    sys.exit(0)


if __name__ == "__main__":
    URI = "ws://localhost:3000/Messages"

    # Handle Ctrl-C
    signal.signal(signal.SIGINT, shutdown)
    signal.signal(signal.SIGTERM, shutdown)

    asyncio.run(websocket_loop(URI))

