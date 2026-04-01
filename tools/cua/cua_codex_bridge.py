"""
cua_codex_bridge.py
-------------------
Reads cua_last_event.txt and xlam_session_status.txt and formats
a structured prompt that can be pasted into the Codex chat or
injected via the VS Code terminal.

Usage:
    python tools/cua/cua_codex_bridge.py

Output:
    Prints a ready-to-paste Codex prompt to stdout.
    Also writes it to tools/cua/codex_prompt_draft.txt.
"""

import os
import json
import datetime

ROOT          = os.path.join(os.path.dirname(__file__), "../..")
STATUS_FILE   = os.path.join(ROOT, "xlam_session_status.txt")
EVENT_FILE    = os.path.join(ROOT, "cua_last_event.txt")
OUTPUT_FILE   = os.path.join(os.path.dirname(__file__), "codex_prompt_draft.txt")


def read_status() -> dict:
    if not os.path.exists(STATUS_FILE):
        return {}
    with open(STATUS_FILE, "r", encoding="utf-8") as f:
        content = f.read().strip()
    try:
        return json.loads(content)
    except json.JSONDecodeError:
        return {"raw": content}


def read_event() -> str:
    if not os.path.exists(EVENT_FILE):
        return "(no CUA event file found)"
    with open(EVENT_FILE, "r", encoding="utf-8") as f:
        return f.read().strip()


def build_codex_prompt(status: dict, event: str) -> str:
    ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    error   = status.get("error", "unknown")
    btn     = status.get("button_clicked", "unknown")
    ev_ts   = status.get("timestamp", ts)

    prompt = f"""## CUA Session Handoff — {ts}

### What happened
An Excel VBA runtime error modal was detected and auto-dismissed by the CUA watcher.

- **Error message:** `{error}`
- **Button clicked:** `{btn}`
- **Dismissed at:** `{ev_ts}`
- **Excel status:** unblocked and ready

### Raw CUA event
{event}

### Your task
Please resume the invSys XLAM build from where it left off.
1. Review the error above — it may indicate a null object reference or missing
   worksheet/range in the last macro that ran.
2. Check `build_xlam_run.log` for context on which module was executing.
3. Apply the minimum fix, re-run the affected module, and confirm the XLAM
   builds cleanly.
4. Update `xlam_session_status.txt` with `{{"status": "build_complete"}}` when done.
"""
    return prompt


def main():
    status = read_status()
    event  = read_event()
    prompt = build_codex_prompt(status, event)

    print(prompt)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        f.write(prompt)
    print(f"\n[bridge] Prompt also saved to: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
