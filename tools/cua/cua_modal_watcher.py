"""
cua_modal_watcher.py
--------------------
CUA (Computer-Using Agent) watcher that:
  1. Takes periodic screenshots
  2. Uses OpenAI vision to detect Excel VBE error modals
  3. Clicks 'End' to dismiss the modal
  4. Reads the VBE code editor highlights BEFORE dismissing:
       YELLOW highlight = the line where execution stopped (the fault line)
       BLUE highlight   = any lines with breakpoints set
  5. Logs the error + highlighted code lines to xlam_session_status.txt
  6. Writes cua_last_event.txt for cua_codex_bridge.py to format into a
     ready-to-paste Codex prompt

VBE highlight color conventions:
  YELLOW = current execution point — the line VBE stopped on (runtime error)
  BLUE   = breakpoint lines — lines the developer marked for debugging

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
    Single vision call that checks for ALL of:
      A) An Excel VBA/VBE runtime error modal dialog
      B) The YELLOW-highlighted line in the VBE code editor
         = the line where execution stopped (the fault line)
      C) Any BLUE-highlighted lines in the VBE code editor
         = lines with breakpoints set by the developer

    IMPORTANT: the yellow and blue referred to here are VBE editor
    highlight colors inside the Visual Basic Editor code pane —
    NOT anything in VS Code or Codex.

    Returns:
    {
      "modal_detected":  bool,
      "error_text":      str,      # full error message from the dialog
      "button_label":    str,      # 'End' | 'Debug' | 'OK'
      "vbe_visible":     bool,     # is the VBE editor window open/visible?
      "yellow_line":     str,      # exact code text of the yellow-highlighted line
      "yellow_line_num": str,      # line number if readable, else null
      "blue_lines":      [str],    # list of exact code text of blue-highlighted lines
    }
    """
    prompt = (
        "You are monitoring a Windows desktop running Microsoft Excel with the "
        "Visual Basic Editor (VBE) open.\n\n"

        "From this screenshot, extract THREE things:\n\n"

        "1. EXCEL ERROR MODAL: Is there a VBA/VBE runtime error dialog box visible?\n"
        "   - If yes: what is the exact error message text?\n"
        "   - If yes: which button to click safely: 'End', 'Debug', or 'OK'?\n\n"

        "2. VBE YELLOW HIGHLIGHT (execution stopped line):\n"
        "   In the VBE code editor pane, look for a line highlighted in YELLOW.\n"
        "   This is the line where VBA execution stopped due to the error.\n"
        "   - What is the exact code text on that yellow line?\n"
        "   - What is the line number, if visible in the margin?\n"
        "   If no yellow line is visible, return null.\n\n"

        "3. VBE BLUE HIGHLIGHTS (breakpoint lines):\n"
        "   In the VBE code editor pane, look for lines highlighted in BLUE.\n"
        "   These are breakpoints the developer has set.\n"
        "   - List the exact code text of every blue-highlighted line.\n"
        "   If no blue lines are visible, return an empty list [].\n\n"

        "NOTE: The yellow and blue highlights are inside the VBE code editor window —"
        " the Visual Basic Editor that opens when you press Alt+F11 in Excel.\n\n"

        "Respond ONLY as valid JSON with these exact keys:\n"
        '{"modal_detected": true/false, "error_text": "...", "button_label": "End", '
        '"vbe_visible": true/false, "yellow_line": "...", "yellow_line_num": "42", '
        '"blue_lines": ["...", "..."]}'
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
        max_tokens=400,
    )
    raw = response.choices[0].message.content.strip()
    if raw.startswith("```"):
        raw = raw.split("\n", 1)[-1].rsplit("```", 1)[0].strip()
    return json.loads(raw)


def get_button_coords(label: str, b64_image: str) -> tuple:
    """Ask vision for the pixel coordinates of a named dialog button."""
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

def update_status_file(result: dict, button_clicked: str):
    """
    Write the full event snapshot to xlam_session_status.txt.
    Includes the error dialog text AND the VBE highlight context
    (yellow fault line + blue breakpoint lines) so Codex knows
    exactly which code line caused the error.
    """
    status = {
        "timestamp":      datetime.datetime.now().isoformat(),
        "event":          "vba_modal_dismissed",
        "error":          result.get("error_text", "unknown"),
        "button_clicked": button_clicked,
        "status":         "excel_unblocked",
        "vbe_highlights": {
            # YELLOW = line where execution stopped (the fault)
            "yellow_fault_line":     result.get("yellow_line"),
            "yellow_fault_line_num": result.get("yellow_line_num"),
            # BLUE = breakpoint lines the developer set
            "blue_breakpoint_lines": result.get("blue_lines", []),
        },
    }
    with open(STATUS_FILE, "w", encoding="utf-8") as f:
        json.dump(status, f, indent=2)
    print(f"[CUA] Status file updated: {STATUS_FILE}")


def append_to_log(result: dict, button_clicked: str):
    """Append a timestamped entry to build_xlam_run.log."""
    ts          = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    yellow_line = result.get("yellow_line") or "(not captured)"
    yellow_num  = result.get("yellow_line_num") or "?"
    blue_lines  = result.get("blue_lines") or []
    blue_fmt    = "\n    ".join(blue_lines) if blue_lines else "(none)"

    entry = (
        f"\n[{ts}] CUA_WATCHER: VBA modal dismissed via '{button_clicked}'\n"
        f"  Error:                    {result.get('error_text', 'unknown')}\n"
        f"  Yellow line (fault) [{yellow_num}]: {yellow_line}\n"
        f"  Blue lines (breakpoints): {blue_fmt}\n"
    )
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(entry)


def notify_codex(result: dict, button_clicked: str):
    """Write cua_last_event.txt for cua_codex_bridge.py to consume."""
    yellow_line = result.get("yellow_line") or "(not captured)"
    yellow_num  = result.get("yellow_line_num") or "?"
    blue_lines  = result.get("blue_lines") or []
    blue_fmt    = "\n  ".join(blue_lines) if blue_lines else "(none)"

    msg = (
        f"[CUA] Excel VBA modal dismissed via '{button_clicked}'.\n"
        f"Error: {result.get('error_text', 'unknown')}\n\n"
        f"VBE Yellow highlight (fault line [{yellow_num}]):\n"
        f"  {yellow_line}\n\n"
        f"VBE Blue highlights (breakpoints):\n"
        f"  {blue_fmt}\n\n"
        f"xlam_session_status.txt updated. Codex may now resume."
    )
    with open(EVENT_FILE, "w", encoding="utf-8") as f:
        f.write(msg)
    print(f"[CUA] Codex notification written to cua_last_event.txt")


# ── Main Loop ─────────────────────────────────────────────────────────────────

def main():
    print("[CUA] Modal watcher started.")
    print("[CUA] Watching for: Excel VBE error modals | yellow fault line | blue breakpoints")
    print(f"[CUA] Poll interval: {POLL_INTERVAL}s | Status file: {STATUS_FILE}")
    print()

    while True:
        try:
            b64    = screenshot_as_base64()
            result = ask_vision_full(b64)

            if result.get("modal_detected"):
                error_text   = result.get("error_text", "unknown error")
                button_label = result.get("button_label", "End")
                yellow_line  = result.get("yellow_line") or "(not captured)"
                yellow_num   = result.get("yellow_line_num") or "?"
                blue_lines   = result.get("blue_lines") or []

                print(f"[CUA] *** MODAL DETECTED ***")
                print(f"[CUA]   Error:              {error_text!r}")
                print(f"[CUA]   Yellow line [{yellow_num}]:  {yellow_line}")
                if blue_lines:
                    print(f"[CUA]   Blue lines:")
                    for bl in blue_lines:
                        print(f"[CUA]     - {bl}")
                print(f"[CUA]   Clicking '{button_label}'...")

                clicked = find_and_click_button(button_label)
                label_used = button_label if clicked else "FAILED"

                if clicked:
                    print(f"[CUA]   '{button_label}' clicked successfully.")
                else:
                    print(f"[CUA]   WARNING: Could not locate '{button_label}' button.")

                update_status_file(result, label_used)
                append_to_log(result, label_used)
                notify_codex(result, label_used)

                time.sleep(3)  # let Excel recover before next poll
            else:
                vbe_status = "VBE open" if result.get("vbe_visible") else "VBE not visible"
                print(f"[CUA] OK — no modal | {vbe_status}", end="\r")

        except json.JSONDecodeError as e:
            print(f"[CUA] JSON parse error from vision model: {e}")
        except Exception as e:
            print(f"[CUA] Unexpected error: {e}")

        time.sleep(POLL_INTERVAL)


if __name__ == "__main__":
    main()
