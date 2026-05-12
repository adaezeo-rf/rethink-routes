# Rethink Food — Route Generator

A Flask web app that turns a weekly member spreadsheet into optimised delivery manifests for NYC meal delivery operations. Deployed on Vercel.

---

## What it does

1. **Upload** an `.xlsx` member list (exported from the member database)
2. **Geocode** each address via Google Maps API (parallel, ~10s for any batch)
3. **Assign** members to one of 14 routes (Mon–Fri, across Bronx, Brooklyn, Manhattan, Queens) based on ZIP code
4. **Optimise** each route using nearest-neighbour + 2-opt TSP from the Tribeca depot
5. **Output** per-route manifests (`.xlsx` + `.csv`), a kitchen packing list, and an interactive map

---

## File structure

```
api/
  index.py              # Flask app — all routes, generation logic, download endpoints
templates/
  base.html             # Dark-theme layout shell + shared CSS
  login.html            # Password auth page
  index.html            # Landing / upload page
  results.html          # Results: maps, stop lists, downloads
rethink_routes.py       # Core routing logic — route definitions, geocoding, TSP optimisation
geocode_cache.json      # Persistent address → lat/lon cache (committed to repo)
requirements.txt        # Python dependencies
vercel.json             # Vercel serverless config (maxDuration: 300s)
.env.example            # Required environment variables
routing_overview.html   # Standalone visual explainer for operations managers
```

---

## Local setup

```bash
# 1. Clone
git clone https://github.com/adaezeo-rf/rethink-routes.git
cd rethink-routes

# 2. Install dependencies
pip install -r requirements.txt

# 3. Copy and fill in environment variables
cp .env.example .env
# Edit .env with your values (see Environment Variables below)

# 4. Run
flask --app api/index.py run
# App is at http://localhost:5000
```

---

## Environment variables

| Variable | Required | Description |
|---|---|---|
| `APP_PASSWORD` | Yes | Password shown on the login screen |
| `SECRET_KEY` | Yes | Flask session signing key — any random string |
| `GOOGLE_MAPS_API_KEY` | Recommended | Enables parallel geocoding (~10s). Without it the app falls back to sequential Nominatim (~2s/address, prone to timeouts) |

On Vercel these are set in **Project Settings → Environment Variables**. Locally they live in a `.env` file (never commit this).

---

## Deployment (Vercel)

The app is deployed as a single Python serverless function at `api/index.py`.

```bash
# Deploy via Vercel CLI
vercel --prod
```

Or push to `master` — Vercel auto-deploys on every push.

**Key config (`vercel.json`):**
- `maxDuration: 300` — 5-minute function timeout (requires Vercel Pro)
- All routes are proxied to `api/index.py`

---

## Routing logic

### Routes
14 routes defined in `rethink_routes.py` → `ROUTES` list. Each entry is:
```python
("letter", "display_name", "borough", "day", ["zip1", "zip2", ...])
```

### ZIP assignment
- Members are assigned to routes by ZIP code match
- ZIPs that appear in two routes are **load-balanced** (fewest-stops-first)
- Load balancing also enforces a **2-day minimum gap** between deliveries to the same household address
- Hard overrides in `ZIP_OVERRIDES` (e.g. Ridgewood → Brooklyn route)

### Geocode cache
`geocode_cache.json` is committed to the repo and loaded at startup. On Vercel the filesystem is read-only, so new geocodes are cached in memory for that warm instance only. To grow the cache permanently: run locally, let it geocode, then commit the updated `geocode_cache.json`.

### Address overrides
For addresses the geocoder gets wrong, add entries to `ADDRESS_OVERRIDES` in `rethink_routes.py`:
```python
ADDRESS_OVERRIDES = {
    ("1 river place", "10036"): (40.7589, -74.0023),  # lat, lon
}
```
Key is `(addr1.strip().lower(), zipcode)`.

### Optimisation
Each route runs: **nearest-neighbour** from the Tribeca depot → **2-opt** improvement, with depot endpoints fixed. The optimised distance (including depot legs) is shown on the results page.

---

## Spreadsheet format

The uploaded `.xlsx` must have these column headers:

| Column | Notes |
|---|---|
| `Member ID` | Unique identifier |
| `Status` | Only rows with `Active` are processed |
| `Box Size` | Large / Medium / Small / Four-Date |
| `Address Line 1` | Street address |
| `Address Line 2` | Apt/unit (optional) |
| `City` / `State` / `Zip` | |
| `Phone Number` | |
| `Delivery Instructions` | |
| `Available Delivery Days` | |
| `Meal Preferences/Allergens` | Optional — surfaced as flags |

---

## Adding a new ZIP code or route

Edit `ROUTES` in `rethink_routes.py`. Each route is a tuple:
```python
("X", "Day_Borough_Name", "Borough", "Weekday", ["zip1", "zip2"])
```
No other changes needed — the app picks it up automatically on the next generation.

---

## Key dependencies

| Package | Purpose |
|---|---|
| `flask` | Web framework |
| `googlemaps` | Parallel geocoding (Google Maps API) |
| `geopy` | Nominatim fallback geocoder |
| `openpyxl` | Read member spreadsheets, write XLSX manifests |
| `pandas` | (available, minimal use) |
| `folium` | Legacy map builder (superseded by Leaflet.js in the UI) |
