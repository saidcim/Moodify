# Moodify

A Spotify bot that rebuilds one playlist for me every two days, based on what I
actually listened to. Runs on GitHub Actions, no server, costs nothing.

I used AI for multiple parts of the project, but the ideas, architectural design, and
choices that influenced the project's progress were all my own.

## Main problem

An AI that picks music has no idea if you liked what it picked. Ask a model to rate its
own playlist and it gives itself an 8 out of 10 forever.

So it counts plays instead. A track carried over from last time either got played again
or it did not, and that is not an opinion.

Spotify only gives you your last 50 plays though, which over 48 hours undercounts
badly. So a second workflow runs every hour doing nothing but adding to the counter.

## How it works

    every 2 days:  history -> mood -> candidates -> curator -> playlist
                   then measure last cycle and tune the strategy

    every hour:    recently_played -> deduplicate -> counter

State lives in a secret Gist, not this repo, since it holds my whole listening history.
Each run pulls it at the start and pushes it at the end. That is all `gist_state.py`
does.

### The carry-over spiral

Some tracks get carried over so the playlist is not all strangers. My first version
asked an AI how many, but the code rejected every increase, so the number could only
fall. Few plays, lower carry_over, nothing familiar left, even fewer plays. It hit zero
and stayed there for four cycles.

Now it compares the carried tracks against the fresh ones in the same playlist. Carried
got played more, +1. Less, -1. Floor of 2 so it cannot bottom out again. No AI call.

### Two scores that must not be compared

There is the model's opinion of its own playlist, and engagement, from real play counts.

I mixed them once. Engagement only started at cycle 12, so the old cycles were 8/10 AI
scores and the new ones 2-4/10 engagement scores. Improvement came out as -5.97 while
the bot was actually getting better, and it kept writing "you are getting worse" into
its own prompt.

### Free tier

Groq allows about 8000 tokens a minute and three calls back to back blew past it, so
there is a rolling window that waits when a request will not fit. Prompts got smaller
too, track lists are just "Title - Artist" now instead of full dictionaries.

When Groq retired `llama-3.3-70b-versatile` the bot stopped dead because I had the model
name hardcoded in three places. It walks a chain now.

### Silent failure

Four runs in August were green with no cycle produced. The poll was swallowing its
exceptions, so broken auth looked exactly like success. Both commands exit non-zero now.

## Demo

There is a working demo in `demo/`, served straight from GitHub Pages. It runs on made
up data in `demo_state.json`, is not wired to any Spotify account, and the buttons do
nothing. Turn on Pages for the repo and it is at `/demo/`.

## Dashboard

`index.html` is a static page that reads the Gist and draws it. Cycle history, mood log,
the strategy it settled on and why, and its own score against the real one.

There is a text box on it. Whatever I type goes into the next prompt above everything
else. "That was too gloomy, make the next one more upbeat" works.

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
