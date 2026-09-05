import os
import sys
import glob
import json
import time
import logging
import datetime
import threading
import openpyxl
import spotipy
from spotipy.oauth2 import SpotifyOAuth
from groq import Groq

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler()],
)
log = logging.getLogger(__name__)


def _require_env(name: str) -> str:
    val = os.environ.get(name)
    if not val:
        raise SystemExit(
            f"ERROR: required environment variable is not set: {name}. "
            f"On GitHub Actions add it under repository Secrets; "
            f"for local runs see env.example."
        )
    return val


SPOTIFY_CLIENT_ID = _require_env("SPOTIFY_CLIENT_ID")
SPOTIFY_CLIENT_SECRET = _require_env("SPOTIFY_CLIENT_SECRET")
SPOTIFY_REDIRECT_URI = os.environ.get("SPOTIFY_REDIRECT_URI", "http://127.0.0.1:5000/callback")
SPOTIFY_REFRESH_TOKEN = _require_env("SPOTIFY_REFRESH_TOKEN")
GROQ_API_KEY = _require_env("GROQ_API_KEY")
USER_NOTE = os.environ.get("USER_NOTE", "").strip()[:500]
PLAYLIST_NAME = os.environ.get("PLAYLIST_NAME", "AI Daily Mix")
PLAYLIST_SIZE = int(os.environ.get("PLAYLIST_SIZE", "40"))
DISCOVERY_QUOTA = min(int(os.environ.get("DISCOVERY_QUOTA", "8")), max(PLAYLIST_SIZE // 3, 1))

DATA_DIR = os.environ.get("DATA_DIR", "/data" if os.path.isdir("/data") else ".")
os.makedirs(DATA_DIR, exist_ok=True)
HISTORY_FILE_PREFIX = "playlist_history_cycle_"
STATE_FILE = os.path.join(DATA_DIR, "bot_state.json")
DASHBOARD_FILE = os.path.join(DATA_DIR, "dashboard_state.json")


def history_file_for_cycle(cycle: int) -> str:
    return os.path.join(DATA_DIR, f"{HISTORY_FILE_PREFIX}{int(cycle):03d}.xlsx")


def _latest_history_file() -> str | None:
    files = glob.glob(os.path.join(DATA_DIR, f"{HISTORY_FILE_PREFIX}*.xlsx"))
    if not files:
        return None

    def _cycle_num(p):
        try:
            return int(os.path.basename(p)[len(HISTORY_FILE_PREFIX):-len(".xlsx")])
        except ValueError:
            return -1

    return max(files, key=_cycle_num)


ARCHIVE_LIMIT = 50
USER_NOTE_WINDOW = 3
USER_NOTE_MAX_CHARS = 200

GROQ_MODEL_CHAIN = [m for m in [
    os.environ.get("GROQ_MODEL", ""),
    "openai/gpt-oss-120b",
    "llama-3.3-70b-versatile",
    "openai/gpt-oss-20b",
    "llama-3.1-8b-instant",
] if m]
_ACTIVE_GROQ_MODEL: str | None = None

GROQ_TPM_BUDGET = int(os.environ.get("GROQ_TPM_BUDGET", "7000"))
_GROQ_USAGE: list = []

DEFAULT_CARRY_OVER = 5
CARRY_OVER_MIN, CARRY_OVER_MAX = 0, 10
CARRY_OVER_FLOOR = 2
CARRY_TUNE_DEADBAND = 0.10

is_running = False
STATE_LOCK = threading.Lock()

AFFINITY_DECAY = 0.95
AFFINITY_PRUNE_BELOW = 0.05