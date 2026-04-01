"""
cua_modal_watcher.py
--------------------
CUA (Computer-Using Agent) watcher that:
  1. Takes periodic screenshots
  2. Uses OpenAI vision to detect Excel VBE error modals
  3. Captures the yellow (remaining) and blue (used) Codex token indicators
     visible in the VS Code Codex panel on EVERY poll — not just on errors
  4. Clicks 'End' to dismiss any Excel VBA modal
  5. Logs the error + token snapshot to xlam_session_status.txt
  6. Writes cua_last_event.txt for cua_codex_bridge.py to format into a
     ready-to-paste Codex prompt

Token color conventions (Codex panel, top-right counter):
  YELLOW tokens = context tokens remaining  (warning: getting close to limit)
  BLUE   tokens = tokens used so far in current session

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
from io import BytesIO

import pyautogui
from PIL import Image
from openai import OpenAI

# ── Config ────────────────────────────────────────────────────────────────────
STATUS_FILE   = os.path.join(os.path.dirname(__file__), "../../xlam_session_status.txt")
LOG_FILE      = os.path.join(os.path.dirname(__file__), "../../build_xlam_run.log")
EVENT_FILE    = os.path.join(os.path.dirname(__file__), "../../cua_last_event.txt")
POLL_INTERVAL = 2          # seconds between screenshots
MODEL         = "gpt-4o"   # vision-capable model
CLIENT        = OpenAI()   # uses OPENAI_API_KEY from environment


# ── Vision Helpers ────────────────────────────────────────────────────────────

def screenshot_as_base64() -> str:
    """Capture full screen and return as base64-encoded PNG string."""
    img = pyautogui.screenshot()
    buf = BytesIO()
    img.save(buf, format="PNG")
    return base64.b64encode(buf.getvalue()).decode()


def ask_vision_full(b64_image: str) -> dict:
    """
    Single vision call that checks for BOTH:
      A) An Excel VBA/VBE runtime error modal
      B) The yellow (remaining) and blue (used) Codex token counters
         visible in the VS Code Codex chat panel

    Returns a dict:
    {
      "modal_detected":   bool,
      "error_text":       str,
      "button_label":     str,   # 'End' | 'Debug' | 'OK'
      "tokens_yellow":    str,   # remaining tokens, e.g. "14,203" or null
      "tokens_blue":      str,   # used tokens, e.g. "83,451" or null
    }

    Token color conventions:
      YELLOW = context tokens REMAINING  (warning indicator — close to limit)
      BLUE   = tokens USED so far in this Codex session
    """
    prompt = (
        "You are monitoring a Windows desktop running Excel and VS Code with Codex.\n\n"
        "From this screenshot, extract TWO things:\n\n"
        "1. EXCEL MODAL: Is there an Excel VBA/VBE runtime error dialog visible?\n"
        "   - If yes: what is the exact error message text?\n"
        "   - If yes: which button to click to dismiss it safely: 'End', 'Debug', or 'OK'?\n\n"
        "2. CODEX TOKEN COUNTERS: Look for the Codex token usage display in VS Code.\n"
        "   It typically shows two numbers near the Codex chat input or header area.\n"
        "   - YELLOW number = context tokens REMAINING (warning color, getting low)\n"
        "   - BLUE number   = tokens USED so far in this session\n"
        "   Read both numbers exactly as shown (including commas).\n"
        "   If a counter is not visible, return null for that field.\n\n"
        "Respond ONLY as valid JSON with these exact keys:\n"
        '{"modal_detected": true/false, "error_text": "...", "button_label": "End", '
        '"tokens_yellow": "14,203", "tokens_blue": "83,451"}'
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
        max_tokens=300,
    )
    raw = response.choices[0].message.content.strip()
    if raw.startswith("```"):
        raw = raw.split("\n", 1)[-1].rsplit("```", 1)[0].strip()
    return json.loads(raw)


def get_button_coords(label: str, b64_image: str) -> tuple:
    """Ask vision for the pixel coordinates of a named button. Returns (x, y) or (None, None)."""
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
                     "image_url": {"url": f"data:image/png;base64,{b64_image}"}},
                ],
            }
        ],
        max_tokens=64,
    )
    raw = resp.choices[0].message.content.strip()
    if raw.startswith("```"):
        raw = raw.split("\n", 1)[-1].rsplit("```", 1)[0].strip()
    coords = json.loads(raw)
    return coords.get("x"), coords.get("y")


def find_and_click_button(label: str) -> bool:
    """
    Locate and click a dialog button by label.
    Tries pyautogui image-match first, falls back to vision coordinates.
    Returns True if click was attempted.
    """
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

    b64 = screenshot_as_base64()
    x, y = get_button_coords(label, b64)
    if x and y:
        pyautogui.click(x, y)
        return True
    return False


# ── State Writers ─────────────────────────────────────────────────────────────

def update_status_file(error_text: str, button_clicked: str,
                       tokens_yellow: str, tokens_blue: str):
    """
    Write full session snapshot to xlam_session_status.txt.
    Codex reads this file to know:
      - whether Excel is unblocked
      - what error occurred
      - current token budget (yellow = remaining, blue = used)
    """
    status = {
        "timestamp":      datetime.datetime.now().isoformat(),
        "event":          "vba_modal_dismissed",
        "error":          error_text,
        "button_clicked": button_clicked,
        "status":         "excel_unblocked",
        "tokens": {
            "yellow_remaining": tokens_yellow,   # warning: context tokens left
            "blue_used":        tokens_blue,      # tokens consumed this session
        },
    }
    with open(STATUS_FILE, "w", encoding="utf-8") as f:
        json.dump(status, f, indent=2)
    print(f"[CUA] Status file updated: {STATUS_FILE}")


def update_status_tokens_only(tokens_yellow: str, tokens_blue: str):
    """
    On a clean poll (no modal), only update the token snapshot
    so Codex always has a fresh token budget reading.
    """
    existing = {}
    if os.path.exists(STATUS_FILE):
        try:
            with open(STATUS_FILE, "r", encoding="utf-8") as f:
                existing = json.load(f)
        except Exception:
            pass

    existing["token_snapshot"] = {
        "timestamp":        datetime.datetime.now().isoformat(),
        "yellow_remaining": tokens_yellow,
        "blue_used":        tokens_blue,
    }
    with open(STATUS_FILE, "w", encoding="utf-8") as f:
        json.dump(existing, f, indent=2)


def append_to_log(error_text: str, tokens_yellow: str, tokens_blue: str):
    """Append a timestamped entry to build_xlam_run.log."""
    ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    entry = (
        f"\n[{ts}] CUA_WATCHER: VBA modal dismissed.\n"
        f"  Error:           {error_text}\n"
        f"  Tokens remaining (yellow): {tokens_yellow}\n"
        f"  Tokens used      (blue):   {tokens_blue}\n"
    )
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(entry)


def notify_codex(error_text: str, tokens_yellow: str, tokens_blue: str):
    """Write cua_last_event.txt for cua_codex_bridge.py to read."""
    msg = (
        f"[CUA] Excel VBA modal dismissed.\n"
        f"Error:                     {error_text!r}\n"
        f"Tokens remaining (yellow): {tokens_yellow}\n"
        f"Tokens used      (blue):   {tokens_blue}\n"
        f"xlam_session_status.txt updated. Codex may now resume."
    )
    with open(EVENT_FILE, "w", encoding="utf-8") as f:
        f.write(msg)
    print(f"[CUA] Codex notification written to cua_last_event.txt")


# ── Main Loop ─────────────────────────────────────────────────────────────────

def main():
    print("[CUA] Modal watcher started.")
    print("[CUA] Watching for: Excel VBE error modals | Codex yellow+blue token counters")
    print(f"[CUA] Poll interval: {POLL_INTERVAL}s | Status file: {STATUS_FILE}")
    print()

    while True:
        try:
            b64    = screenshot_as_base64()
            result = ask_vision_full(b64)

            tokens_yellow = result.get("tokens_yellow") or "n/a"
            tokens_blue   = result.get("tokens_blue")   or "n/a"

            if result.get("modal_detected"):
                error_text   = result.get("error_text", "unknown error")
                button_label = result.get("button_label", "End")

                print(f"[CUA] *** MODAL DETECTED ***")
                print(f"[CUA]   Error:             {error_text!r}")
                print(f"[CUA]   Token remaining:   {tokens_yellow}  (yellow)")
                print(f"[CUA]   Tokens used:       {tokens_blue}   (blue)")
                print(f"[CUA]   Clicking '{button_label}'...")

                clicked = find_and_click_button(button_label)
                if clicked:
                    print(f"[CUA]   '{button_label}' clicked successfully.")
                else:
                    print(f"[CUA]   WARNING: Could not locate '{button_label}' button.")

                update_status_file(
                    error_text,
                    button_label if clicked else "FAILED",
                    tokens_yellow,
                    tokens_blue,
                )
                append_to_log(error_text, tokens_yellow, tokens_blue)
                notify_codex(error_text, tokens_yellow, tokens_blue)

                time.sleep(3)  # let Excel recover
            else:
                # No modal — still update the token snapshot so Codex stays informed
                update_status_tokens_only(tokens_yellow, tokens_blue)
                print(
                    f"[CUA] OK | yellow(remaining): {tokens_yellow} "
                    f"| blue(used): {tokens_blue}",
                    end="\r"
                )

        except json.JSONDecodeError as e:
            print(f"[CUA] JSON parse error from vision model: {e}")
        except Exception as e:
            print(f"[CUA] Unexpected error: {e}")

        time.sleep(POLL_INTERVAL)


if __name__ == "__main__":
    main()
