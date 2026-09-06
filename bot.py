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


def _default_user_profile() -> dict:
    return {
        "version": 1,
        "last_updated": None,
        "learned_patterns": [],
        "genre_affinity": {},
        "artist_affinity": {},
        "mood_music_map": {},
        "next_advice": "",
        "adaptation_metrics": {
            "cycles_completed": 0,
            "score_basis": "insufficient_engagement_history",
            "avg_score_first_3": 0.0,
            "avg_score_last_3": 0.0,
            "improvement_delta": 0.0,
            "avg_ai_score_last_3": 0.0,
            "avg_plays_trend": "0%",
            "carry_over_success_rate": None,
            "carry_over_measured": {},
            "confidence": 0.0,
        },
    }


def _default_dynamic_config() -> dict:
    return {
        "carry_over": DEFAULT_CARRY_OVER,
        "last_tuned_cycle": 0,
        "tune_reason": "",
    }


def _default_state() -> dict:
    return {
        "playlist_id": os.environ.get("PLAYLIST_ID", ""),
        "last_update": None,
        "cycle": 0,
        "feedback_history": [],
        "mood_history": [],
        "playlist_archive": [],
        "user_profile": _default_user_profile(),
        "dynamic_config": _default_dynamic_config(),
        "ai_notes": "",
        "cumulative_plays": {},
        "seen_play_events": [],
        "user_notes": [],
    }


def _load_state_from_file(path: str) -> dict | None:
    if not os.path.exists(path):
        return None
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        log.warning(f"Could not read state file ({path}): {e}")
        return None


def _merge_state_defaults(state: dict) -> dict:
    default = _default_state()
    for key, val in default.items():
        if key not in state:
            state[key] = val
    if "user_profile" in state:
        profile_default = _default_user_profile()
        for k, v in profile_default.items():
            if k not in state["user_profile"]:
                state["user_profile"][k] = v
        metrics = profile_default["adaptation_metrics"]
        for k, v in metrics.items():
            if k not in state["user_profile"].get("adaptation_metrics", {}):
                state["user_profile"]["adaptation_metrics"][k] = v

    dyn_default = _default_dynamic_config()
    dyn = state.get("dynamic_config") or {}
    dyn.pop("last_recovery_cycle", None)
    for k, v in dyn_default.items():
        if k not in dyn:
            dyn[k] = v
    dyn["carry_over"] = max(CARRY_OVER_MIN, min(CARRY_OVER_MAX, int(dyn["carry_over"])))
    dyn["carry_over"] = min(dyn["carry_over"], PLAYLIST_SIZE)
    state["dynamic_config"] = dyn
    return state


def load_state() -> dict:
    state = _load_state_from_file(STATE_FILE)
    if state is None and STATE_FILE != "bot_state.json":
        state = _load_state_from_file("bot_state.json")
    if state is None:
        log.info("No saved state found, starting clean.")
        state = _default_state()
    else:
        log.info(f"State loaded (cycle #{state.get('cycle', 0)}).")
    return _merge_state_defaults(state)


def write_dashboard_state(state: dict):
    try:
        payload = {
            "generated_at": datetime.datetime.now(datetime.timezone.utc).isoformat(),
            "cycle": state.get("cycle", 0),
            "last_update": state.get("last_update"),
            "playlist_id": state.get("playlist_id"),
            "is_running": False,
            "feedback_history": state.get("feedback_history", [])[-30:],
            "mood_history": state.get("mood_history", [])[-30:],
            "user_profile": state.get("user_profile", {}),
            "dynamic_config": state.get("dynamic_config", {}),
            "cumulative_plays": state.get("cumulative_plays", {}),
            "user_notes": state.get("user_notes", [])[-10:],
            "playlist_archive_summary": state.get("playlist_archive", [])[-10:],
        }
        with open(DASHBOARD_FILE, "w", encoding="utf-8") as f:
            json.dump(payload, f, indent=2, ensure_ascii=False)
    except Exception as e:
        log.warning(f"Could not write dashboard_state.json: {e}")


def save_state(state: dict):
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, indent=2, ensure_ascii=False)
    write_dashboard_state(state)


TRACK_HEADERS = ["Cycle", "Date", "Time", "Spotify ID", "Track", "Artist", "Album",
                 "AI Score", "Play Count"]
SUMMARY_HEADERS = ["Cycle", "Datetime", "Mood", "Energy", "Score", "Track Count",
                   "Avg Plays", "Carried Tracks", "AI Analysis Summary"]


def _get_or_create_sheets(wb):
    if "Tracks" in wb.sheetnames:
        ws_tracks = wb["Tracks"]
    else:
        ws_tracks = wb.active
        ws_tracks.title = "Tracks"
    if "Cycle Summary" in wb.sheetnames:
        ws_summary = wb["Cycle Summary"]
    else:
        ws_summary = wb.create_sheet("Cycle Summary")
    if ws_tracks.max_row == 1 and ws_tracks.cell(1, 1).value is None:
        for col, h in enumerate(TRACK_HEADERS, 1):
            ws_tracks.cell(1, col).value = h
    if ws_summary.max_row == 1 and ws_summary.cell(1, 1).value is None:
        for col, h in enumerate(SUMMARY_HEADERS, 1):
            ws_summary.cell(1, col).value = h
    return ws_tracks, ws_summary


def save_to_excel(tracks: list, cycle: int, score=None, play_counts: dict | None = None,
                  history_file: str | None = None):
    play_counts = play_counts or {}
    path = history_file or history_file_for_cycle(cycle)
    now = datetime.datetime.now()
    date_str = now.strftime("%Y-%m-%d")
    time_str = now.strftime("%H:%M:%S")

    wb = openpyxl.load_workbook(path) if os.path.exists(path) else openpyxl.Workbook()
    ws_tracks, _ = _get_or_create_sheets(wb)

    for track in tracks:
        plays = play_counts.get(track["id"], 0)
        ws_tracks.append([
            cycle, date_str, time_str, track["id"], track["name"],
            track["artist"], track["album"], score if score is not None else "-", plays,
        ])
    wb.save(path)
    log.info(f"Saved {len(tracks)} tracks to Excel (cycle {cycle}) -> {os.path.basename(path)}")


def save_cycle_summary_to_excel(cycle: int, mood_data: dict, score: float, track_count: int,
                                avg_plays: float, carry_count: int, analysis: str,
                                history_file: str | None = None):
    path = history_file or history_file_for_cycle(cycle)
    now = datetime.datetime.now()
    datetime_str = now.strftime("%Y-%m-%d %H:%M:%S")

    wb = openpyxl.load_workbook(path) if os.path.exists(path) else openpyxl.Workbook()
    _, ws_summary = _get_or_create_sheets(wb)
    ws_summary.append([
        cycle, datetime_str,
        mood_data.get("mood", "-"),
        mood_data.get("energy_level", "-"),
        score, track_count, round(avg_plays, 2), carry_count,
        (analysis or "")[:200],
    ])
    wb.save(path)
    log.info(f"Saved cycle summary to Excel (cycle {cycle}) -> {os.path.basename(path)}")


    def get_spotify() -> spotipy.Spotify:
    auth = SpotifyOAuth(
        client_id=SPOTIFY_CLIENT_ID,
        client_secret=SPOTIFY_CLIENT_SECRET,
        redirect_uri=SPOTIFY_REDIRECT_URI,
        scope=(
            "user-read-recently-played user-top-read user-library-read "
            "playlist-modify-public playlist-modify-private playlist-read-private"
        ),
    )
    token_info = auth.refresh_access_token(SPOTIFY_REFRESH_TOKEN)
    return spotipy.Spotify(auth=token_info["access_token"])


def _parse_played_at(raw: str) -> datetime.datetime:
    dt = datetime.datetime.fromisoformat(raw.replace("Z", "+00:00"))
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=datetime.timezone.utc)
    return dt


def get_listening_data(sp: spotipy.Spotify) -> dict:
    data = {}
    recent = sp.current_user_recently_played(limit=50)
    recent_tracks = []
    cutoff = datetime.datetime.now(datetime.timezone.utc) - datetime.timedelta(days=3)
    for item in recent["items"]:
        played_at = _parse_played_at(item["played_at"])
        if played_at >= cutoff:
            t = item["track"]
            recent_tracks.append({
                "id": t["id"], "name": t["name"],
                "artist": t["artists"][0]["name"], "album": t["album"]["name"],
                "played_at": item["played_at"],
            })
    data["recent_tracks"] = recent_tracks

    top_short = sp.current_user_top_tracks(limit=50, time_range="short_term")
    data["top_short"] = [{"id": t["id"], "name": t["name"], "artist": t["artists"][0]["name"],
                          "popularity": t["popularity"]} for t in top_short["items"]]

    top_medium = sp.current_user_top_tracks(limit=50, time_range="medium_term")
    data["top_medium"] = [{"id": t["id"], "name": t["name"], "artist": t["artists"][0]["name"],
                           "popularity": t["popularity"]} for t in top_medium["items"]]

    top_artists = sp.current_user_top_artists(limit=20, time_range="short_term")
    data["top_artists"] = [{"id": a["id"], "name": a["name"], "genres": a["genres"],
                            "popularity": a.get("popularity", 0)} for a in top_artists["items"]]

    saved = sp.current_user_saved_tracks(limit=50)
    data["saved_tracks"] = [{"id": item["track"]["id"], "name": item["track"]["name"],
                             "artist": item["track"]["artists"][0]["name"]} for item in saved["items"]]

    return data


def get_playlist_play_counts(playlist_track_ids, recent_tracks):
    counts = {tid: 0 for tid in playlist_track_ids}
    for t in recent_tracks:
        if t["id"] in counts:
            counts[t["id"]] += 1
    return counts


def build_listening_fingerprint(listening_data: dict) -> dict:
    genres = []
    for a in listening_data.get("top_artists", [])[:10]:
        genres.extend(a.get("genres", [])[:3])
    genre_counts: dict[str, int] = {}
    for g in genres:
        genre_counts[g] = genre_counts.get(g, 0) + 1
    top_genres = sorted(genre_counts, key=genre_counts.get, reverse=True)[:5]
    return {
        "top_artists": [a["name"] for a in listening_data.get("top_artists", [])[:8]],
        "top_genres": top_genres,
        "track_count_3d": len(listening_data.get("recent_tracks", [])),
    }