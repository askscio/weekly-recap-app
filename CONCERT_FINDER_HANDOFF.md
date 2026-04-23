# Concert Finder App — Complete Handoff

## Project Overview
Concert Finder is a web app that helps Billy and Gretchen discover concerts in their favorite vacation destinations by combining:
- **Spotify**: Pulls user's top artists + related artists for discovery
- **Bandsintown**: API for artist tour dates
- **Ticketmaster**: API for venue/city-based events
- **Flask**: Simple web UI for non-technical users to add artists/destinations and run searches

## Current Status (March 28, 2026)
✅ **Core Backend**: Fully working (all clients, matchers, config)
✅ **CLI Tool**: Fully working (main.py with interactive prompts)
✅ **Storage**: Persistent custom_data.json for saving artists/destinations
✅ **Flask Web UI**: Built and Flask installed
✅ **Server**: Running on http://localhost:8000 (in background as of last session)

## File Structure
```
concert-finder/
├── .env                           # API keys (Spotify, Ticketmaster, Bandsintown)
├── .venv/                         # Python virtualenv (venv with Python 3.13.7)
├── requirements.txt               # Dependencies (spotipy, requests, flask, etc.)
├── config.py                      # Hardcoded destinations + date windows
├── storage.py                     # JSON load/save for custom_data.json
├── spotify_client.py              # Spotify OAuth + top artists + similar artists
├── bandsintown_client.py          # Artist event search
├── ticketmaster_client.py         # Destination/city-based event search
├── matcher.py                     # Dedup events, match to user/discovery artists
├── main.py                        # CLI with interactive prompts (--save-custom, --list-custom, etc.)
├── webapp.py                      # Flask app routes (/, /add_artist, /add_destination, /search)
├── templates/
│   ├── index.html                 # Main page (artists, destinations, add forms)
│   └── results.html               # Search results table
├── custom_data.json               # (Auto-created) Saved custom artists/destinations
├── README.md                      # Setup & usage docs
└── results.txt                    # (Auto-created by CLI) Text dump of results
```

## Environment Setup

### Python Environment
- **Type**: venv
- **Version**: Python 3.13.7
- **Location**: `/Users/billy.schuett/weekly-recap-app/concert-finder/.venv`
- **Activate**: `source /Users/billy.schuett/weekly-recap-app/concert-finder/.venv/bin/activate`

### API Keys (.env file)
Already configured and present in the repo:
```
SPOTIFY_CLIENT_ID=1623400a6ae143a4b77c7b58b1c8f830
SPOTIFY_CLIENT_SECRET=e4897ee083fa485fa4d2e9ec3a5168fb
SPOTIFY_REDIRECT_URI=https://localhost:8888/callback
TICKETMASTER_API_KEY=aw8WlBrMHY2BlaASnXuUcp9pMM9jREea
BANDSINTOWN_APP_ID=concert_finder_app
```

### Dependencies (from requirements.txt)
- spotipy (Spotify API)
- requests (HTTP client)
- python-dotenv (load .env)
- tabulate (CLI tables)
- colorama (CLI colors)
- flask (web framework)

## How to Run

### Web UI (Recommended for users)
```bash
cd /Users/billy.schuett/weekly-recap-app/concert-finder
bash -c "/path/to/.venv/bin/python webapp.py"
```
Runs on **http://localhost:8000**

### CLI (Non-interactive)
```bash
cd /Users/billy.schuett/weekly-recap-app/concert-finder
.venv/bin/python main.py --add-artist "Artist Name" --no-interactive
```

### CLI (Interactive with persistent saves)
```bash
cd /Users/billy.schuett/weekly-recap-app/concert-finder
.venv/bin/python main.py --save-custom
```

### List saved custom artists/destinations
```bash
.venv/bin/python main.py --list-custom
```

## Key Features Built

### Web UI Routes
- **GET `/`**: Home page — shows Spotify top artists, saved custom artists, destinations, forms to add both
- **POST `/add_artist`**: Save artist name to custom_data.json
- **POST `/add_destination`**: Save destination (city, state, country) to custom_data.json
- **POST `/search`**: Run full search across all sources, return results.html

### CLI Features
- Interactive prompts for artists/destinations (only if TTY)
- `--add-artist "Name"` flag for non-interactive adds
- `--save-custom` flag to persist interactive additions
- `--list-custom` to view saved data
- `--no-interactive` to disable TTY prompts
- Outputs results to `results.txt` and console tables

### Core Logic
1. **Spotify Auth**: Uses SpotifyOAuth (first run opens browser for consent)
2. **Artist Search**: Gets top artists from 3 time ranges (short/medium/long term) + deduplicates
3. **Discovery**: Fetches related artists for each top artist, tracks "similar_to" relationships
4. **Event Aggregation**: Searches Bandsintown (artist-based) + Ticketmaster (city-based), deduplicates
5. **Matching**: Categorizes events into `your_artist`, `discovery`, `outside_dates` based on travel window (May 25 - Aug 25, 2026 primary; Sep 1 - Dec 31 secondary)

## Known Behavior / Notes

### Spotify OAuth
- First run will attempt browser-based OAuth callback to `https://localhost:8888/callback`
- Uses `.cache` file in concert-finder directory to store refresh token
- Must complete the flow once, then it's cached for future runs

### Date Windows (hardcoded in config.py)
```python
PRIMARY_DATE_WINDOW = ('2026-05-25', '2026-08-25')    # Memorial Day to Aug 25
SECONDARY_DATE_WINDOW = ('2026-09-01', '2026-12-31')  # Sep 1 to end of year
```
Events outside primary window are marked `in_travel_window: false` but still reported

### Destinations (default, can be extended)
Red Rocks, Thompson's Point (Portland ME), Greek Theatre (Berkeley), Savannah, Charleston, Nashville, Austin, Asheville, Montreal, Vancouver

### Custom Data Persistence
- Saved to `custom_data.json` in concert-finder directory
- Format: `{ "artists": [{artist_name: "...", ...}], "destinations": [{name, city, state, country, search_terms}] }`
- Auto-loaded on every run, merged with defaults

### Rate Limiting
- Bandsintown: 0.5s sleep between artist searches
- Ticketmaster: 0.3s sleep between city searches
- No aggressive retry logic; errors are logged and skipped

## Recent Additions (This Session)

### 1. Storage Module (storage.py)
- `load_custom_data()`: Read custom_data.json or return empty dict
- `save_custom_data(data)`: Write dict to custom_data.json

### 2. Updated Main.py
- Now loads custom artists/destinations on startup
- Interactive prompts (TTY-only) for adding artists/destinations
- `--save-custom` flag persists new entries
- `--list-custom` lists saved entries and exits

### 3. Flask Web App (webapp.py)
- Simple routes for home, add artist, add destination, search
- Loads top Spotify artists on page load
- Saves custom entries to disk via storage module
- Passes search results to template

### 4. HTML Templates
- **index.html**: Two-column layout — artists on left, destinations on right, add forms for each
- **results.html**: Tabular view of results — your_artist, discovery, outside_dates

### 5. README.md
- Updated to highlight web UI as primary interface
- CLI instructions as fallback
- Setup steps and API key config

## Common Tasks

### To add a new destination
1. Go to http://localhost:8000
2. Scroll to "Saved" destinations section
3. Fill in City, State/Province, Country
4. Click "Save destination"
5. Now included in search

### To add a custom artist
1. Go to http://localhost:8000
2. Scroll to "Saved Custom Artists"
3. Type artist name
4. Click "Save artist"
5. Now included in search

### To run a search
1. Add/verify artists and destinations
2. Click "Run search" at bottom of page
3. Wait for results (may take 30-60s depending on API response times)
4. View results in tabular format

### To reset custom data
1. Delete `custom_data.json` in concert-finder folder
2. Refresh web page or restart app

## Potential Issues & Fixes

### Flask won't start
- Ensure Flask is installed: `pip install flask`
- Ensure you're in the concert-finder directory when running webapp.py
- Use absolute path to Python: `/path/to/.venv/bin/python webapp.py`

### Spotify auth fails
- Check `.env` has valid SPOTIFY_CLIENT_ID and SECRET
- First OAuth attempt must complete; may need to manually navigate to http://localhost:8888/callback if browser doesn't auto-open
- Delete `.cache` file and try again if stuck

### No results found
- Check date windows in config.py (currently May 25 - Aug 25 2026 primary)
- Verify destination spellings match Bandsintown/Ticketmaster data
- Check API rate limiting — add more sleep between requests if needed

### Custom data not persisting
- Ensure `custom_data.json` is writable in concert-finder directory
- Check file is not corrupted (should be valid JSON)

## Next Steps / Ideas for Enhancement

1. **Web UI Polish**
   - Add real CSS styling (currently minimal)
   - Loading spinner during search
   - Delete buttons for custom artists/destinations
   - Filter results by date/artist type

2. **Search Optimization**
   - Cache API responses to avoid rate limit delays
   - Async search (separate thread or Celery) so UI doesn't hang
   - Progress bar showing search stages

3. **Testing & CI**
   - Unit tests for matcher.py, storage.py, client modules
   - Integration test with mock API responses
   - GitHub Actions CI

4. **Advanced Features**
   - Price range display (Ticketmaster has pricing)
   - Notifications when new shows added for an artist
   - Save favorite shows / build personal itinerary
   - Multi-user support (currently single-user)

5. **Deployment**
   - Docker container for easy setup
   - Deploy to Heroku, AWS Lambda, or Vercel
   - Use environment variables for API keys (not hardcoded .env)

## Quick Test Command
```bash
cd /Users/billy.schuett/weekly-recap-app/concert-finder && \
bash -c "/Users/billy.schuett/weekly-recap-app/concert-finder/.venv/bin/python -c 'import webapp, storage, main, matcher, spotify_client, bandsintown_client, ticketmaster_client, config; print(\"ALL_IMPORTS_OK\")'"
```
If this prints `ALL_IMPORTS_OK`, the environment is ready.

## Contact & Context
- User: Billy
- Date built: March 28, 2026
- Purpose: Find concerts in vacation destinations for Billy & Gretchen
- Tech: Python 3.13 + Flask + Spotify/Ticketmaster/Bandsintown APIs
