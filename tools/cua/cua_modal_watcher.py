"""
cua_modal_watcher.py
--------------------
CUA (Computer-Using Agent) watcher that:
  1. Takes periodic screenshots
  2. Uses OpenAI vision to detect Excel VBE error modals
  3. Clicks 'End' to dismiss the modal
  4. Logs the error + token context to xlam_session_status.txt
  5. Sends a summary message to Codex via VS Code terminal

Requirements:
    pip install openai pillow pyautogui pygetwindow

Usage:
    python tools/cua/cua_modal_watcher.py
"""

import os
import time
import json
import base64
import datetime
import subprocess
from io import BytesIO

import pyautogui
import pygetwindow as gw
from PIL import Image
from openai import OpenAI

# ── Config ────────────────────────────────────────────────────────────────────
STATUS_FILE   = os.path.join(os.path.dirname(__file__), "../../xlam_session_status.txt")
LOG_FILE      = os.path.join(os.path.dirname(__file__), "../../build_xlam_run.log")
POLL_INTERVAL = 2          # seconds between screenshots
MODEL         = "gpt-4o"   # vision-capable model
CLIENT        = OpenAI()   # uses OPENAI_API_KEY from environment

# ── Helpers ───────────────────────────────────────────────────────────────────

def screenshot_as_base64() -> str:
    """Capture full screen and return as base64-encoded PNG string."""
    img = pyautogui.screenshot()
    buf = BytesIO()
    img.save(buf, format="PNG")
    return base64.b64encode(buf.getvalue()).decode()


def ask_vision(b64_image: str) -> dict:
    """
    Ask GPT-4o whether an Excel VBE error modal is visible.
    Returns a dict: { "modal_detected": bool, "error_text": str, "button_label": str }
    """
    prompt = (
        "You are monitoring a Windows desktop. "
        "Look at this screenshot and tell me:\n"
        "1. Is there an Excel VBA/VBE runtime error dialog box visible? (yes/no)\n"
        "2. If yes, what is the exact error message text?\n"
        "3. If yes, what button should be clicked to dismiss it safely — 'End', 'Debug', or 'OK'?\n"
        "Respond ONLY as valid JSON: "
        '{"modal_detected": true/false, "error_text": "...", "button_label": "End"}'
    )
    response = CLIENT.chat.completions.create(
        model=MODEL,
        messages=[
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": prompt},
                    {"type": "image_url",
                     "image_url": {"url": f"data:image/png;base64,{b64_image}"}},
                ],
            }
        ],
        max_tokens=256,
    )
    raw = response.choices[0].message.content.strip()
    # Strip markdown code fences if present
    if raw.startswith("```"):
        raw = raw.split("\n", 1)[-1].rsplit("```", 1)[0].strip()
    return json.loads(raw)


def find_and_click_button(label: str) -> bool:
    """
    Locate a button on screen by text using vision-located coordinates.
    Falls back to pyautogui.locateOnScreen for simple cases.
    Returns True if click was attempted.
    """
    # Try image-based locate first (fast path)
    try:
        btn = pyautogui.locateOnScreen(
            f"tools/cua/assets/btn_{label.lower()}.png",
            confidence=0.8
        )
        if btn:
            pyautogui.click(pyautogui.center(btn))
            return True
    except Exception:
        pass

    # Fallback: ask vision for coordinates
    b64 = screenshot_as_base64()
    coord_prompt = (
        f"In this screenshot, find the '{label}' button in the Excel error dialog. "
        "Return ONLY JSON: {\"x\": <pixel_x>, \"y\": <pixel_y>} for the button center. "
        "If not found, return {\"x\": null, \"y\": null}."
    )
    resp = CLIENT.chat.completions.create(
        model=MODEL,
        messages=[
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": coord_prompt},
                    {"type": "image_url",
                     "image_url": {"url": f"data:image/png;base64,{b64}"}},
                ],
            }
        ],
        max_tokens=64,
    )
    raw = resp.choices[0].message.content.strip()
    if raw.startswith("```"):
        raw = raw.split("\n", 1)[-1].rsplit("```", 1)[0].strip()
    coords = json.loads(raw)
    if coords.get("x") and coords.get("y"):
        pyautogui.click(coords["x"], coords["y"])
        return True
    return False


def update_status_file(error_text: str, button_clicked: str):
    """Write dismissal event to xlam_session_status.txt for Codex to read."""
    status = {
        "timestamp": datetime.datetime.now().isoformat(),
        "event": "vba_modal_dismissed",
        "error": error_text,
        "button_clicked": button_clicked,
        "status": "excel_unblocked",
    }
    with open(STATUS_FILE, "w", encoding="utf-8") as f:
        json.dump(status, f, indent=2)
    print(f"[CUA] Status file updated: {STATUS_FILE}")


def append_to_log(error_text: str):
    """Append a timestamped entry to build_xlam_run.log."""
    ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    entry = f"\n[{ts}] CUA_WATCHER: VBA modal dismissed. Error was: {error_text}\n"
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(entry)


def notify_codex(error_text: str):
    """
    Open a new VS Code terminal and echo the error context so Codex
    can pick it up from the terminal history.
    Requires VS Code CLI ('code') to be in PATH.
    """
    msg = (
        f"[CUA] Excel VBA modal dismissed. "
        f"Error was: {error_text!r}. "
        f"xlam_session_status.txt updated. Codex may now resume."
    )
    try:
        # Write to a temp file so Codex can read it without terminal dependency
        tmp = os.path.join(os.path.dirname(__file__), "../../cua_last_event.txt")
        with open(tmp, "w") as f:
            f.write(msg)
        print(f"[CUA] Codex notification written to cua_last_event.txt")
    except Exception as e:
        print(f"[CUA] Could not write Codex notification: {e}")


# ── Main Loop ─────────────────────────────────────────────────────────────────

def main():
    print("[CUA] Modal watcher started. Watching for Excel VBE error dialogs...")
    print(f"[CUA] Poll interval: {POLL_INTERVAL}s | Status file: {STATUS_FILE}")

    while True:
        try:
            b64 = screenshot_as_base64()
            result = ask_vision(b64)

            if result.get("modal_detected"):
                error_text   = result.get("error_text", "unknown error")
                button_label = result.get("button_label", "End")

                print(f"[CUA] Modal detected! Error: {error_text!r}")
                print(f"[CUA] Clicking '{button_label}'...")

                clicked = find_and_click_button(button_label)
                if clicked:
                    print(f"[CUA] '{button_label}' clicked successfully.")
                else:
                    print(f"[CUA] WARNING: Could not locate '{button_label}' button.")

                update_status_file(error_text, button_label if clicked else "FAILED")
                append_to_log(error_text)
                notify_codex(error_text)

                # Extra pause to let Excel recover before next poll
                time.sleep(3)
            else:
                print(f"[CUA] No modal. Sleeping {POLL_INTERVAL}s...", end="\r")

        except json.JSONDecodeError as e:
            print(f"[CUA] JSON parse error from vision model: {e}")
        except Exception as e:
            print(f"[CUA] Unexpected error: {e}")

        time.sleep(POLL_INTERVAL)


if __name__ == "__main__":
    main()
