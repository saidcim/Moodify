import os
import sys
import requests

STATE_FILENAME = "bot_state.json"
DATA_DIR = os.environ.get("DATA_DIR", ".")
LOCAL_PATH = os.path.join(DATA_DIR, STATE_FILENAME)

GIST_ID = os.environ.get("GIST_ID", "")
GIST_TOKEN = os.environ.get("GIST_TOKEN", "")
API_URL = f"https://api.github.com/gists/{GIST_ID}"
HEADERS = {
    "Authorization": f"Bearer {GIST_TOKEN}",
    "Accept": "application/vnd.github+json",
    "X-GitHub-Api-Version": "2022-11-28",
}


def _require_config():
    if not GIST_ID or not GIST_TOKEN:
        sys.exit("ERROR: GIST_ID and GIST_TOKEN environment variables must be set.")


def pull():
    _require_config()
    resp = requests.get(API_URL, headers=HEADERS, timeout=30)
    resp.raise_for_status()
    files = resp.json().get("files", {}) or {}
    entry = files.get(STATE_FILENAME)
    if not entry:
        print("No state in the gist; the bot will start clean.")
        return
    content = entry.get("content", "")
    if entry.get("truncated") and entry.get("raw_url"):
        raw = requests.get(entry["raw_url"], headers=HEADERS, timeout=30)
        raw.raise_for_status()
        content = raw.text
    os.makedirs(DATA_DIR, exist_ok=True)
    with open(LOCAL_PATH, "w", encoding="utf-8") as f:
        f.write(content)
    print(f"State downloaded from the gist ({len(content)} bytes) -> {LOCAL_PATH}")


def push():
    _require_config()
    if not os.path.exists(LOCAL_PATH):
        print("No local bot_state.json; skipping push.")
        return
    with open(LOCAL_PATH, encoding="utf-8") as f:
        content = f.read()
    payload = {"files": {STATE_FILENAME: {"content": content}}}
    resp = requests.patch(API_URL, headers=HEADERS, json=payload, timeout=30)
    resp.raise_for_status()
    print(f"State uploaded to the gist ({len(content)} bytes).")


if __name__ == "__main__":
    cmd = sys.argv[1].lower() if len(sys.argv) > 1 else ""
    if cmd == "pull":
        pull()
    elif cmd == "push":
        push()
    else:
        sys.exit("Usage: python gist_state.py [pull|push]")
