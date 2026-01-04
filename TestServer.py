# -*- coding: utf-8 -*-
"""
TestServer for WinToTalk (modern websockets)
- Sends 6 sample Say messages on SPACE press with delay
- Ctrl-C exits cleanly
- Accepts any client (simulate /Messages)
"""

import asyncio
import json
import threading
import signal
import sys
import msvcrt
import websockets
from queue import Queue

PORT = 3000
DELAY_BETWEEN_MESSAGES = 4  # seconds

payloads = [
    {"Language": "English", "Voice": {"Name": "None"}, "Gender": "None", "Payload": "This is english Neutral test.", "Speaker": "Test"},
    {"Language": "English", "Voice": {"Name": "Female"}, "Gender": "Female", "Payload": "This is english Female test.", "Speaker": "Test"},
    {"Language": "English", "Voice": {"Name": "Male"}, "Gender": "Male", "Payload": "This english Male test.", "Speaker": "Test"},
    {"Language": "German", "Voice": {"Name": "None"}, "Gender": "None", "Payload": "Das ist deutsch neutral test.", "Speaker": "Test"},
    {"Language": "German", "Voice": {"Name": "Female"}, "Gender": "Female", "Payload": "Das ist deutsch Weiblich test.", "Speaker": "Test"},
    {"Language": "German", "Voice": {"Name": "Male"}, "Gender": "Male", "Payload": "Das ist deutsch männlich test.", "Speaker": "Test"},
]

clients = set()
send_queue = Queue()
stop_event = threading.Event()

# ------------------------
# WebSocket handler
# ------------------------
async def handler(ws):
    clients.add(ws)
    print(f"[TestServer] Client connected: {ws.remote_address}")
    try:
        while not stop_event.is_set():
            await asyncio.sleep(0.1)
    except asyncio.CancelledError:
        pass
    finally:
        clients.remove(ws)
        print(f"[TestServer] Client disconnected: {ws.remote_address}")

# ------------------------
# Send payloads from queue
# ------------------------
async def send_loop():
    while not stop_event.is_set():
        try:
            task = send_queue.get_nowait()
        except:
            await asyncio.sleep(0.1)
            continue
        # task is just a trigger to send all 6 payloads
        for i, payload in enumerate(payloads, 1):
            msg = json.dumps({"Type": "Say", **payload})
            for ws in list(clients):
                try:
                    await ws.send(msg)
                except Exception as e:
                    print(f"[TestServer] Send error: {e}")
            print(f"[TestServer] Sent {i}/6: {payload['Payload']}")
            await asyncio.sleep(DELAY_BETWEEN_MESSAGES)

# ------------------------
# Keyboard thread
# ------------------------
def keyboard_thread():
    print("[TestServer] Press SPACE to send test messages, Ctrl-C to quit")
    while not stop_event.is_set():
        if msvcrt.kbhit():
            key = msvcrt.getch()
            if key == b" ":
                send_queue.put("send")  # trigger send
            elif key == b"\x03":  # Ctrl-C
                stop_event.set()
                break
        else:
            # sleep to reduce CPU usage
            import time
            time.sleep(0.05)

# ------------------------
# Main server
# ------------------------
async def main():
    async with websockets.serve(handler, "localhost", PORT):
        print(f"[TestServer] Running WebSocket server on ws://localhost:{PORT}")
        threading.Thread(target=keyboard_thread, daemon=True).start()
        await send_loop()
    print("[TestServer] Shutting down...")

if __name__ == "__main__":
    signal.signal(signal.SIGINT, lambda s, f: stop_event.set())
    signal.signal(signal.SIGTERM, lambda s, f: stop_event.set())
    asyncio.run(main())

