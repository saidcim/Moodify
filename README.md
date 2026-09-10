# Moodify

A Spotify bot that rebuilds one playlist for me every two days, based on what I
actually listened to.

## How it works
    every 2 days:  history -> mood -> candidates -> curator -> playlist
                   then measure last cycle and tune the strategy

    every hour:    recently_played -> deduplicate -> counter

State lives in a secret Gist, not this repo, since it holds my whole listening history.
Each run pulls it at the start and pushes it at the end. That is all `gist_state.py`
does.

## Demo

There is a working demo in `demo/`, served straight from GitHub Pages. It runs on made
up data in `demo_state.json`, is not wired to any Spotify account, and the buttons do
nothing. Turn on Pages for the repo and it is at `/demo/`.

## Setup

Fork it, the cycle workflow commits back to the repo.

**Spotify app** at developer.spotify.com/dashboard. Redirect URI exactly
`http://127.0.0.1:5000/callback`, the IP, not localhost. Tick Web API.

**Refresh token.** Save as `get_token.py` and run:

```python
import json
from spotipy.oauth2 import SpotifyOAuth

SCOPE = ("user-read-recently-played user-top-read user-library-read "
         "playlist-modify-public playlist-modify-private playlist-read-private")

auth = SpotifyOAuth(
    client_id="YOUR_CLIENT_ID",
    client_secret="YOUR_CLIENT_SECRET",
    redirect_uri="http://127.0.0.1:5000/callback",
    scope=SCOPE,
    cache_path=".spotify_cache",
)
auth.get_access_token(as_dict=False)
print(json.load(open(".spotify_cache"))["refresh_token"])
```

A browser opens, you approve, you land on a page that fails to load. Copy the whole URL
from the address bar and paste it back into the terminal. Delete the script after.

**Groq key** from console.groq.com.

**Gist.** A secret gist with one file `bot_state.json` containing `{}`. The id is in the
URL. Then a classic token with only the `gist` scope, fine grained ones cannot touch
gists.

**Secrets.** Settings, Secrets and variables, Actions. Six of them:
`SPOTIFY_CLIENT_ID`, `SPOTIFY_CLIENT_SECRET`, `SPOTIFY_REFRESH_TOKEN`, `GROQ_API_KEY`,
`GIST_ID`, `GIST_TOKEN`.

**Write permission.** Settings, Actions, General, Workflow permissions, Read and write.
Without it the commit step 403s.

**Playlist name.** In both workflow files, and they have to match:

    PLAYLIST_NAME: "Your Playlist Name"
    PLAYLIST_SIZE: "20"

Then Actions, Moodify Cycle, Run workflow.

For the dashboard, point `GIST_ID` and `BOT_REPO` in `index.html` at your own and turn
on Pages. Keep in mind a secret gist is not a private one, so publishing a working panel
means anyone with the URL can read your listening stats. Leave the placeholder in if you
only want the demo public.

## Locally

```bash
pip install -r requirements.txt
cp env.example .env
python bot.py cycle
python bot.py poll
```

Settings are all environment variables, see `env.example`.

## When it breaks

Green workflow but no cycle, poll saying "No new plays." — token is missing scopes.

`No Groq model was usable` — bad or rate limited key, the log says what it tried.

Short playlist — not much history to choose from yet.

"Connection error" on the dashboard — wrong gist id.

## Licence

Do whatever you want with it.
