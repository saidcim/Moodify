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