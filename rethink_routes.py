"""
rethink_routes.py
-----------------
Weekly route generator for Rethink Food deliveries.

Usage:
    python rethink_routes.py                        # uses default Excel path
    python rethink_routes.py "path/to/members.xlsx" # custom file

Outputs (written to a dated folder inside Downloads):
    Route_<X>_<Name>_map.html        — interactive optimized map per route
    Route_<X>_<Name>_manifest.csv    — ordered stop list for driver
    Kitchen_Packing_List.csv         — box counts + allergens per route
    Flags.txt                        — anomalies to review before delivery
    geocode_cache.json               — persistent address cache (reused each week)
"""

import csv
import io
import json
import math
import re
import sys
import time
import warnings
from datetime import date
from pathlib import Path

# Force UTF-8 output on Windows so Unicode chars don't crash the console
if sys.stdout.encoding != "utf-8":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

import folium
import openpyxl
from folium.features import DivIcon
from geopy.exc import GeocoderServiceError, GeocoderTimedOut
from geopy.geocoders import Nominatim

warnings.filterwarnings("ignore", category=UserWarning)

# ── Configuration ─────────────────────────────────────────────────────────────

CACHE_FILE = Path(__file__).parent / "geocode_cache.json"

# Stop-count limits per route
MAX_STOPS_HARD  = 25    # routes above this are flagged
MAX_STOPS_SOFT  = 28    # allowed when consecutive stops are densely clustered
CLOSE_STOP_MILES = 0.15 # ~2 city blocks — threshold for "close" stops

# Depot addresses — all routes start here
DEPOT_START = {
    "addr1":   "10 Desbrosses St",
    "zipcode": "10013",
    "city":    "New York",
    "state":   "NY",
    "borough": "Manhattan",
    "label":   "START — Tribeca Rooftop",
}
# Bronx routes return here
DEPOT_BRONX_END = {
    "addr1":   "1955 Turnbull Ave",
    "zipcode": "10473",
    "city":    "Bronx",
    "state":   "NY",
    "borough": "Bronx",
    "label":   "END — Bronx Return Depot",
}
# All other routes return here
DEPOT_OTHER_END = {
    "addr1":   "630 Flushing Ave",
    "zipcode": "11206",
    "city":    "Brooklyn",
    "state":   "NY",
    "borough": "Brooklyn",
    "label":   "END — Brooklyn Return Depot",
}

# 14 routes: (letter, display_name, borough_hint, day, [zip codes])
#
# Some zip codes intentionally appear in two routes (e.g. 11433 in both
# Tue_Jamaica and Fri_Jamaica). Members in overlapping zips are distributed
# by load-balancing — whichever matching route has fewer stops gets the next
# member — so routes stay near equal size across days.
ROUTES = [
    ("A", "Mon_Bronx_North",          "Bronx",     "Monday",    [
        "10463","10466","10467","10469","10475",
    ]),
    ("B", "Mon_Brooklyn",             "Brooklyn",  "Monday",    [
        "11201","11206","11211","11216","11217","11218","11219","11220",
        "11221","11231","11237","11238","11249",
    ]),
    ("C", "Mon_Manhattan",            "Manhattan", "Monday",    [
        "10027","10030","10031","10032","10033","10034","10039","10040",
    ]),
    ("D", "Tue_Bronx",                "Bronx",     "Tuesday",   [
        "10452","10453","10468",
    ]),
    ("E", "Tue_Jamaica",              "Queens",    "Tuesday",   [
        "11412","11433","11435","11453",
    ]),
    ("F", "Tue_Queens",               "Queens",    "Tuesday",   [
        "11101","11103","11106","11355","11356","11365","11368","11369",
        "11370","11372","11377","11379","11385",
    ]),
    ("G", "Wed_Brooklyn",             "Brooklyn",  "Wednesday", [
        "11203","11207","11208","11212","11233","11236","11239",
    ]),
    ("H", "Wed_Jamaica_FarRockaways", "Queens",    "Wednesday", [
        "11416","11417","11418","11419","11420","11434","11691","11692",
    ]),
    ("I", "Wed_Manhattan",            "Manhattan", "Wednesday", [
        "10001","10011","10023","10024","10025","10036",
    ]),
    ("J", "Thu_Bronx",                "Bronx",     "Thursday",  [
        "10451","10454","10455","10456",
    ]),
    ("K", "Thu_Manhattan",            "Manhattan", "Thursday",  [
        "10002","10009","10029","10035","10044","10075","10129",
    ]),
    ("L", "Fri_Bronx",                "Bronx",     "Friday",    [
        "10459","10460","10461","10462","10472","10473",
    ]),
    ("M", "Fri_Brooklyn",             "Brooklyn",  "Friday",    [
        "11209","11210","11213","11214","11219","11220","11223",
        "11224","11226","11228","11235",
    ]),
    ("N", "Fri_Jamaica",              "Queens",    "Friday",    [
        "11423","11429","11432","11433","11434","11435","11436",
    ]),
]

BOX_COLORS = {
    "Large":     "#e74c3c",
    "Medium":    "#e67e22",
    "Small":     "#2980b9",
    "Four-Date": "#8e44ad",
    "Unknown":   "#7f8c8d",
}

MAX_ROUTE_MILES = 25.0

ZIP_OVERRIDES = {
    "11385": "B",  # Ridgewood — better with Brooklyn route (near parking lot)
}

# ── Borough bounding boxes for geocode validation ────────────────────────────

BOROUGH_BOUNDS = {
    "Manhattan": {"lat": (40.70, 40.88), "lon": (-74.02, -73.90)},
    "Bronx":     {"lat": (40.78, 40.92), "lon": (-73.93, -73.75)},
    "Brooklyn":  {"lat": (40.57, 40.74), "lon": (-74.05, -73.83)},
    "Queens":    {"lat": (40.54, 40.80), "lon": (-73.96, -73.70)},
}
NYC_BOUNDS = {"lat": (40.49, 40.92), "lon": (-74.26, -73.68)}

# ── Helpers ───────────────────────────────────────────────────────────────────

def clean_box(raw):
    if not raw:
        return "Unknown"
    s = str(raw).strip().lower()
    if s.startswith("l"): return "Large"
    if s.startswith("m"): return "Medium"
    if s.startswith("s"): return "Small"
    if "four" in s or "4-date" in s or "4 date" in s:
        return "Four-Date"
    return "Unknown"


def haversine_miles(a, b):
    R = 3958.8
    lat1, lon1 = math.radians(a[0]), math.radians(a[1])
    lat2, lon2 = math.radians(b[0]), math.radians(b[1])
    dlat, dlon = lat2 - lat1, lon2 - lon1
    h = math.sin(dlat / 2) ** 2 + math.cos(lat1) * math.cos(lat2) * math.sin(dlon / 2) ** 2
    return R * 2 * math.asin(math.sqrt(h))


def route_distance(order, stops):
    return sum(
        haversine_miles(stops[order[i]]["latlon"], stops[order[i + 1]]["latlon"])
        for i in range(len(order) - 1)
    )


def check_stop_limit(ordered_stops):
    """
    Returns a warning string if the route exceeds the stop limit, else None.
    - <= MAX_STOPS_HARD (25): OK
    - MAX_STOPS_HARD < n <= MAX_STOPS_SOFT (28): allowed if stops are dense
    - > MAX_STOPS_SOFT: always flagged
    """
    n = len(ordered_stops)
    if n <= MAX_STOPS_HARD:
        return None

    geocoded = [s for s in ordered_stops if s.get("latlon")]
    if n <= MAX_STOPS_SOFT and len(geocoded) >= 2:
        close = sum(
            1 for i in range(len(geocoded) - 1)
            if haversine_miles(geocoded[i]["latlon"], geocoded[i + 1]["latlon"]) <= CLOSE_STOP_MILES
        )
        close_ratio = close / (len(geocoded) - 1)
        if close_ratio >= 0.75:
            return (
                f"{n} stops (above {MAX_STOPS_HARD} limit). "
                f"Stops are densely clustered — permitted up to {MAX_STOPS_SOFT}."
            )
        return (
            f"{n} stops exceeds the {MAX_STOPS_HARD}-stop limit. "
            f"Stops are not all closely clustered — consider splitting this route."
        )

    return (
        f"{n} stops exceeds maximum capacity of {MAX_STOPS_SOFT}. "
        f"This route must be split before delivery."
    )


# ── Route optimization ────────────────────────────────────────────────────────

def nearest_neighbor(stops):
    """Greedy NN starting from stop index 0. Returns list of indices."""
    n = len(stops)
    unvisited = set(range(1, n))
    route = [0]
    while unvisited:
        last = route[-1]
        nearest = min(unvisited, key=lambda j: haversine_miles(stops[last]["latlon"], stops[j]["latlon"]))
        route.append(nearest)
        unvisited.remove(nearest)
    return route


def nearest_neighbor_from_latlon(stops, start_latlon):
    """Greedy NN starting from an arbitrary lat/lon (not a stop). Returns list of indices."""
    n = len(stops)
    unvisited = set(range(n))
    route = []
    current = start_latlon
    while unvisited:
        nearest = min(unvisited, key=lambda j: haversine_miles(current, stops[j]["latlon"]))
        route.append(nearest)
        current = stops[nearest]["latlon"]
        unvisited.remove(nearest)
    return route


def two_opt(route, stops):
    """
    2-opt improvement. route is a list of indices into stops.
    Index 0 and the last index are kept fixed (depot anchors when used with depots).
    """
    n = len(route)
    improved = True
    while improved:
        improved = False
        for i in range(1, n - 1):
            for j in range(i + 1, n):
                if j - i == 1:
                    continue
                if j == n - 1:
                    if haversine_miles(stops[route[i - 1]]["latlon"], stops[route[j - 1]]["latlon"]) \
                       < haversine_miles(stops[route[i - 1]]["latlon"], stops[route[i]]["latlon"]) - 1e-10:
                        route[i:j] = route[i:j][::-1]
                        improved = True
                    continue
                cur = (haversine_miles(stops[route[i - 1]]["latlon"], stops[route[i]]["latlon"]) +
                       haversine_miles(stops[route[j - 1]]["latlon"], stops[route[j]]["latlon"]))
                new = (haversine_miles(stops[route[i - 1]]["latlon"], stops[route[j - 1]]["latlon"]) +
                       haversine_miles(stops[route[i]]["latlon"],     stops[route[j]]["latlon"]))
                if new < cur - 1e-10:
                    route[i:j] = route[i:j][::-1]
                    improved = True
    return route


def optimize_route(stops, depot_start_latlon=None, depot_end_latlon=None):
    """
    Return stops reordered by nearest-neighbor + 2-opt.

    When depot_start_latlon / depot_end_latlon are provided the NN starts from
    the depot and the 2-opt keeps both depot endpoints fixed. The returned
    ordered list contains only the delivery stops (depots are not included).
    The returned opt_dist includes the depot legs.
    """
    if len(stops) < 2:
        return stops, 0.0, 0.0

    n = len(stops)
    orig_dist = route_distance(list(range(n)), stops)

    # Nearest-neighbor phase
    if depot_start_latlon:
        nn_indices = nearest_neighbor_from_latlon(stops, depot_start_latlon)
    else:
        nn_indices = nearest_neighbor(stops)

    nn_stops = [stops[i] for i in nn_indices]

    # 2-opt phase — augment with depot anchors when available
    if depot_start_latlon and depot_end_latlon:
        aug = [{"latlon": depot_start_latlon}] + nn_stops + [{"latlon": depot_end_latlon}]
        aug_route = list(range(len(aug)))
        opt_aug = two_opt(aug_route, aug)
        ordered = [aug[i] for i in opt_aug[1:-1]]   # strip depot entries
    else:
        opt_idx = two_opt(list(range(len(nn_stops))), nn_stops)
        ordered = [nn_stops[i] for i in opt_idx]

    # Compute optimized distance (including depot legs if present)
    opt_dist = 0.0
    if depot_start_latlon and ordered:
        opt_dist += haversine_miles(depot_start_latlon, ordered[0]["latlon"])
    for i in range(len(ordered) - 1):
        opt_dist += haversine_miles(ordered[i]["latlon"], ordered[i + 1]["latlon"])
    if depot_end_latlon and ordered:
        opt_dist += haversine_miles(ordered[-1]["latlon"], depot_end_latlon)

    return ordered, orig_dist, opt_dist


# ── Geocoding ─────────────────────────────────────────────────────────────────

def load_cache():
    if CACHE_FILE.exists():
        with open(CACHE_FILE, "r") as f:
            raw = json.load(f)
        return {tuple(k.split("|||")): tuple(v) if v else None for k, v in raw.items()}
    return {}


def save_cache(cache):
    raw = {"|||".join(k): list(v) if v else None for k, v in cache.items()}
    with open(CACHE_FILE, "w") as f:
        json.dump(raw, f, indent=2)


def validate_geocode(latlon, borough=None):
    """
    Check that latlon falls within the expected NYC borough bounding box.
    Returns True if valid, False otherwise.
    """
    if latlon is None:
        return False
    lat, lon = latlon
    bounds = BOROUGH_BOUNDS.get(borough, NYC_BOUNDS) if borough else NYC_BOUNDS
    return (bounds["lat"][0] <= lat <= bounds["lat"][1] and
            bounds["lon"][0] <= lon <= bounds["lon"][1])


def make_queries(addr1, zipcode, borough):
    base = addr1.split(",")[0].strip() if "," in addr1 else addr1
    borough_map = {
        "Queens":    "Queens, NY",
        "Bronx":     "Bronx, NY",
        "Manhattan": "New York, NY",
        "Brooklyn":  "Brooklyn, NY",
    }
    city = borough_map.get(borough, "New York, NY")
    return [
        f"{base}, {city} {zipcode}",
        f"{base}, {city}",
        f"{base}, New York, NY {zipcode}",
        f"{base}, New York, NY",
    ]


def geocode_stop(geolocator, addr1, zipcode, borough, cache):
    key = (addr1, zipcode)
    if key in cache:
        return cache[key]

    for query in make_queries(addr1, zipcode, borough):
        try:
            loc = geolocator.geocode(query, timeout=10)
            if loc:
                result = (loc.latitude, loc.longitude)
                if validate_geocode(result, borough):
                    cache[key] = result
                    save_cache(cache)
                    return result
        except (GeocoderTimedOut, GeocoderServiceError):
            pass
        time.sleep(1)

    cache[key] = None
    save_cache(cache)
    return None


# ── Map builder ───────────────────────────────────────────────────────────────

def build_map(ordered_stops, route_letter, route_name, opt_dist,
              day=None, depot_start=None, depot_end=None):
    """
    Build a Folium map for a single route.
    depot_start / depot_end are dicts with keys: latlon, label.
    """
    if not ordered_stops:
        return None

    lats = [s["latlon"][0] for s in ordered_stops]
    lons = [s["latlon"][1] for s in ordered_stops]

    # Include depot coords in center calculation if available
    all_lats = list(lats)
    all_lons = list(lons)
    if depot_start and depot_start.get("latlon"):
        all_lats.append(depot_start["latlon"][0])
        all_lons.append(depot_start["latlon"][1])
    if depot_end and depot_end.get("latlon"):
        all_lats.append(depot_end["latlon"][0])
        all_lons.append(depot_end["latlon"][1])

    center = [sum(all_lats) / len(all_lats), sum(all_lons) / len(all_lons)]
    m = folium.Map(location=center, zoom_start=13, tiles="OpenStreetMap")

    # Title
    day_label = f" &bull; {day}" if day else ""
    title = (
        f'<div style="position:fixed;top:12px;left:50%;transform:translateX(-50%);'
        f'z-index:1000;background:white;padding:8px 20px;border-radius:6px;'
        f'border:1px solid #ccc;box-shadow:0 2px 8px rgba(0,0,0,.3);'
        f'font-family:Arial,sans-serif;font-size:15px;font-weight:bold;">'
        f'Route {route_letter}{day_label} &mdash; {route_name.replace("_", " ")} '
        f'({len(ordered_stops)} stops &bull; {opt_dist:.1f} mi)</div>'
    )
    m.get_root().html.add_child(folium.Element(title))

    # Polyline — include depots at start and end
    polyline_coords = []
    if depot_start and depot_start.get("latlon"):
        polyline_coords.append(depot_start["latlon"])
    polyline_coords.extend(s["latlon"] for s in ordered_stops)
    if depot_end and depot_end.get("latlon"):
        polyline_coords.append(depot_end["latlon"])

    folium.PolyLine(
        locations=polyline_coords, color="#2c3e50", weight=2.5, opacity=0.75,
        tooltip=f"Route {route_letter}"
    ).add_to(m)

    # Depot start marker
    if depot_start and depot_start.get("latlon"):
        depot_html = (
            '<div style="background:#27ae60;color:white;border-radius:4px;'
            'width:32px;height:32px;display:flex;align-items:center;'
            'justify-content:center;font-size:13px;font-weight:bold;'
            'border:2px solid white;box-shadow:0 1px 4px rgba(0,0,0,.5);'
            'font-family:Arial,sans-serif;">S</div>'
        )
        folium.Marker(
            location=depot_start["latlon"],
            popup=folium.Popup(
                f'<b>Route Start</b><br>{depot_start["label"]}<br>'
                f'{depot_start.get("addr1","")}, {depot_start.get("zipcode","")}',
                max_width=260
            ),
            tooltip=depot_start["label"],
            icon=DivIcon(html=depot_html, icon_size=(32, 32), icon_anchor=(16, 16)),
        ).add_to(m)

    # Delivery stop markers
    for i, stop in enumerate(ordered_stops, start=1):
        color  = BOX_COLORS.get(stop["box_size"], "#7f8c8d")
        lat, lon = stop["latlon"]
        border = "3px solid #f1c40f" if stop["flag"] else "2px solid white"

        icon_html = (
            f'<div style="background-color:{color};color:white;border-radius:50%;'
            f'width:28px;height:28px;display:flex;align-items:center;'
            f'justify-content:center;font-size:12px;font-weight:bold;'
            f'border:{border};box-shadow:0 1px 4px rgba(0,0,0,.5);'
            f'font-family:Arial,sans-serif;">{i}</div>'
        )

        allergens = stop["allergens"] if stop["allergens"] and stop["allergens"].lower() not in ("none", "") else "None"
        notes     = stop["notes"]     if stop["notes"]     and stop["notes"].lower()     not in ("none", "") else "N/A"
        di        = stop["delivery_instructions"] if stop["delivery_instructions"] else "N/A"
        flag_html = f'<br><b style="color:#e74c3c;">FLAG: {stop["flag"]}</b>' if stop["flag"] else ""

        popup_html = (
            f'<div style="font-family:Arial,sans-serif;font-size:13px;min-width:240px;">'
            f'<b>Stop #{i}</b>{flag_html}<br>'
            f'<b>Member ID:</b> {stop["member_id"]}<br>'
            f'<b>Address:</b> {stop["display_addr"]}<br>'
            f'<b>Phone:</b> {stop["phone"]}<br>'
            f'<b>Box Size:</b> {stop["box_size"]}<br>'
            f'<b>Allergens:</b> {allergens}<br>'
            f'<b>Delivery Instructions:</b> {di}<br>'
            f'<b>Avail. Days:</b> {stop["available_days"]}<br>'
            f'<b>Notes:</b> {notes}</div>'
        )

        folium.Marker(
            location=[lat, lon],
            popup=folium.Popup(popup_html, max_width=320),
            tooltip=f"Stop #{i} — {stop['box_size']} — {stop['addr1']}",
            icon=DivIcon(html=icon_html, icon_size=(28, 28), icon_anchor=(14, 14)),
        ).add_to(m)

    # Depot end marker
    if depot_end and depot_end.get("latlon"):
        depot_end_html = (
            '<div style="background:#c0392b;color:white;border-radius:4px;'
            'width:32px;height:32px;display:flex;align-items:center;'
            'justify-content:center;font-size:13px;font-weight:bold;'
            'border:2px solid white;box-shadow:0 1px 4px rgba(0,0,0,.5);'
            'font-family:Arial,sans-serif;">E</div>'
        )
        folium.Marker(
            location=depot_end["latlon"],
            popup=folium.Popup(
                f'<b>Route End</b><br>{depot_end["label"]}<br>'
                f'{depot_end.get("addr1","")}, {depot_end.get("zipcode","")}',
                max_width=260
            ),
            tooltip=depot_end["label"],
            icon=DivIcon(html=depot_end_html, icon_size=(32, 32), icon_anchor=(16, 16)),
        ).add_to(m)

    # Legend
    legend_html = (
        '<div style="position:fixed;bottom:30px;right:30px;z-index:1000;'
        'background:white;padding:12px 16px;border-radius:8px;'
        'border:1px solid #ccc;box-shadow:0 2px 8px rgba(0,0,0,.25);'
        'font-family:Arial,sans-serif;font-size:13px;line-height:2;">'
        '<b style="font-size:14px;">Box Size</b><br>'
        '<span style="display:inline-block;width:14px;height:14px;border-radius:50%;'
        'background:#e74c3c;vertical-align:middle;margin-right:6px;"></span>Large<br>'
        '<span style="display:inline-block;width:14px;height:14px;border-radius:50%;'
        'background:#e67e22;vertical-align:middle;margin-right:6px;"></span>Medium<br>'
        '<span style="display:inline-block;width:14px;height:14px;border-radius:50%;'
        'background:#2980b9;vertical-align:middle;margin-right:6px;"></span>Small<br>'
        '<hr style="margin:4px 0;">'
        '<span style="display:inline-block;width:14px;height:14px;border-radius:50%;'
        'border:3px solid #f1c40f;background:#7f8c8d;vertical-align:middle;margin-right:6px;"></span>'
        'Flagged stop<br>'
        '<span style="display:inline-block;width:14px;height:14px;border-radius:4px;'
        'background:#27ae60;vertical-align:middle;margin-right:6px;"></span>Start depot<br>'
        '<span style="display:inline-block;width:14px;height:14px;border-radius:4px;'
        'background:#c0392b;vertical-align:middle;margin-right:6px;"></span>End depot'
        '</div>'
    )
    m.get_root().html.add_child(folium.Element(legend_html))

    return m


# ── Household clustering ─────────────────────────────────────────────────────

_APT_PATTERN = re.compile(
    r",?\s*(?:apt|unit|floor|suite|#)\s*\S*",
    re.IGNORECASE,
)


def detect_household_clusters(stops):
    """Group stops that share the same street address (ignoring apt/unit info).

    Returns a dict mapping stop index (0-based) to a group letter (A, B, C, ...)
    for groups of size >= 2.  Stops not in a cluster are omitted.
    """
    # Normalize: strip apt/unit/floor info, lowercase, strip whitespace
    def _normalize(addr1):
        cleaned = _APT_PATTERN.sub("", addr1)
        return cleaned.strip().lower()

    # Group by normalized address
    groups = {}  # normalized_addr -> [index, ...]
    for i, s in enumerate(stops):
        key = _normalize(s["addr1"])
        groups.setdefault(key, []).append(i)

    # Assign letters to groups of size >= 2
    result = {}
    letter_idx = 0
    for indices in groups.values():
        if len(indices) >= 2:
            letter = chr(ord("A") + letter_idx)
            for idx in indices:
                result[idx] = letter
            letter_idx += 1

    return result


# ── Manifest CSV ──────────────────────────────────────────────────────────────

MANIFEST_HEADERS = [
    "Stop", "Member ID", "Address", "Apt/Unit", "City", "Zip",
    "Phone", "Box Size", "Allergens", "Delivery Instructions",
    "Available Days", "Notes", "Flag",
]


def write_manifest(ordered_stops, path):
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(MANIFEST_HEADERS)
        for i, s in enumerate(ordered_stops, start=1):
            writer.writerow([
                i,
                s["member_id"],
                s["addr1"],
                s["addr2"],
                s["city"],
                s["zipcode"],
                s["phone"],
                s["box_size"],
                s["allergens"],
                s["delivery_instructions"],
                s["available_days"],
                s["notes"],
                s["flag"],
            ])


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    DOWNLOADS  = Path.home() / "Downloads"
    EXCEL_PATH = Path(
        sys.argv[1] if len(sys.argv) > 1
        else r"C:\Users\adaez\Downloads\MTM MEMBERS LIST, Rethink Food..xlsx"
    )
    OUTPUT_DIR = DOWNLOADS / f"RethinkRoutes_{date.today().isoformat()}"

    print(f"\n{'='*60}")
    print(f"  Rethink Food — Route Generator")
    print(f"  Week of {date.today().isoformat()}")
    print(f"{'='*60}\n")

    if not EXCEL_PATH.exists():
        sys.exit(f"ERROR: File not found: {EXCEL_PATH}")

    wb = openpyxl.load_workbook(EXCEL_PATH, read_only=True)
    ws = wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    headers = rows[0]
    col = {h: i for i, h in enumerate(headers) if h is not None}

    required = ["Member ID", "Box Size", "Address Line 1", "City", "State", "Zip",
                "Phone Number", "Delivery Instructions", "Status"]
    missing = [c for c in required if c not in col]
    if missing:
        sys.exit(f"ERROR: Missing columns in spreadsheet: {missing}")

    # Parse all active members
    all_stops, flags_global = [], []

    for row in rows[1:]:
        status = str(row[col["Status"]] or "").strip().lower()
        if status != "active":
            continue

        zipcode = str(row[col["Zip"]] or "").replace(".0", "").strip().zfill(5)
        addr1   = str(row[col["Address Line 1"]] or "").strip()
        addr2   = str(row[col["Address Line 2"]] or "").strip() if "Address Line 2" in col else ""
        city    = str(row[col["City"]]   or "").strip()
        state   = str(row[col["State"]]  or "").strip()

        display_addr = addr1
        if addr2 and addr2.lower() not in ("none", ""):
            display_addr += f", {addr2}"
        display_addr += f", {city}, {state} {zipcode}"

        allergens = str(row[col.get("Meal Preferences/Allergens", -1)] or "").strip() \
            if col.get("Meal Preferences/Allergens") is not None else ""
        avail = str(row[col.get("Available Delivery Days", -1)] or "").strip() \
            if col.get("Available Delivery Days") is not None else ""
        notes_raw = " | ".join(filter(None, [
            str(row[col[c]] or "").strip()
            for c in ["Unnamed: 13", "Unnamed: 14", "Unnamed: 15"] if c in col
        ]))

        flag = ""
        notes_lower = notes_raw.lower()
        if "cancel" in notes_lower:
            flag = "Notes say 'cancel' but status is Active — verify"
        elif "hold" in notes_lower:
            flag = "Notes mention 'hold' — verify"
        if allergens and allergens.lower() not in ("none", ""):
            flag = (flag + " | " if flag else "") + f"Allergen: {allergens}"

        member_id = str(row[col["Member ID"]] or "").replace(".0", "")
        stop = {
            "member_id":              member_id,
            "addr1":                  addr1,
            "addr2":                  addr2,
            "city":                   city,
            "state":                  state,
            "zipcode":                zipcode,
            "display_addr":           display_addr,
            "phone":                  str(row[col["Phone Number"]] or "").strip(),
            "box_size":               clean_box(row[col["Box Size"]]),
            "allergens":              allergens,
            "delivery_instructions":  str(row[col["Delivery Instructions"]] or "").strip(),
            "available_days":         avail,
            "notes":                  notes_raw,
            "flag":                   flag,
            "latlon":                 None,
        }
        all_stops.append(stop)
        if flag:
            flags_global.append(f"Member {member_id} ({display_addr}): {flag}")

    print(f"Active members found: {len(all_stops)}")

    # Build zip -> routes mapping (supports overlapping zips)
    zip_to_routes = {}
    for letter, name, borough, day, zips in ROUTES:
        for z in zips:
            zip_to_routes.setdefault(z, []).append((letter, name, borough, day))

    route_stops = {letter: [] for letter, *_ in ROUTES}
    for stop in all_stops:
        matching = zip_to_routes.get(stop["zipcode"], [])
        if not matching:
            flags_global.append(
                f"Member {stop['member_id']} (zip {stop['zipcode']}) "
                "not assigned to any route — add zip to ROUTES in rethink_routes.py"
            )
        elif len(matching) == 1:
            route_stops[matching[0][0]].append((stop, matching[0][2]))
        else:
            # Load-balance: assign to the matching route with the fewest stops so far
            best = min(matching, key=lambda r: len(route_stops[r[0]]))
            route_stops[best[0]].append((stop, best[2]))

    # Geocode depots
    geocache   = load_cache()
    geolocator = Nominatim(user_agent="rethink_food_router_v2")

    print("Geocoding depots...")
    depot_start_latlon = geocode_stop(
        geolocator, DEPOT_START["addr1"], DEPOT_START["zipcode"], DEPOT_START["borough"], geocache)
    depot_bronx_latlon = geocode_stop(
        geolocator, DEPOT_BRONX_END["addr1"], DEPOT_BRONX_END["zipcode"], DEPOT_BRONX_END["borough"], geocache)
    depot_other_latlon = geocode_stop(
        geolocator, DEPOT_OTHER_END["addr1"], DEPOT_OTHER_END["zipcode"], DEPOT_OTHER_END["borough"], geocache)

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    print(f"Output folder: {OUTPUT_DIR}\n")

    kitchen_rows, summary_lines = [], []

    for letter, name, borough, day, zips in ROUTES:
        entries = route_stops[letter]
        if not entries:
            print(f"Route {letter} ({name}): no active members, skipping.")
            continue

        stops_for_route = [e[0] for e in entries]
        print(f"{'-'*50}")
        print(f"Route {letter} — {name.replace('_',' ')} | {day} | {len(stops_for_route)} members")

        # Geocode stops
        geocoded = []
        for stop in stops_for_route:
            key = (stop["addr1"], stop["zipcode"])
            if key in geocache:
                stop["latlon"] = geocache[key]
            else:
                print(f"  [geocoding] {stop['addr1']} ({stop['zipcode']})...")
                result = geocode_stop(geolocator, stop["addr1"], stop["zipcode"], borough, geocache)
                stop["latlon"] = result
                time.sleep(1)
            if stop["latlon"]:
                geocoded.append(stop)
            else:
                flags_global.append(
                    f"Member {stop['member_id']} ({stop['addr1']}, {stop['zipcode']}): geocoding failed"
                )

        if not geocoded:
            continue

        # Choose end depot
        depot_end_latlon = depot_bronx_latlon if borough == "Bronx" else depot_other_latlon
        depot_end_info   = DEPOT_BRONX_END    if borough == "Bronx" else DEPOT_OTHER_END

        # Optimize with depot anchors
        ordered, orig_dist, opt_dist = optimize_route(geocoded, depot_start_latlon, depot_end_latlon)

        # Stop limit check
        limit_warn = check_stop_limit(ordered)
        if limit_warn:
            flags_global.append(f"Route {letter} ({name.replace('_',' ')}): {limit_warn}")
            print(f"  STOP LIMIT: {limit_warn}")

        saved = orig_dist - (route_distance(list(range(len(geocoded))), geocoded) if orig_dist else 0)
        pct   = ((orig_dist - opt_dist) / orig_dist * 100) if orig_dist > 0 else 0
        print(f"  Distance: {opt_dist:.1f} mi total (incl. depot legs) | {pct:.0f}% improvement")

        # Map
        safe_name = name.replace(" ", "_")
        depot_s   = {**DEPOT_START,    "latlon": depot_start_latlon}
        depot_e   = {**depot_end_info, "latlon": depot_end_latlon}
        m = build_map(ordered, letter, name, opt_dist, day=day, depot_start=depot_s, depot_end=depot_e)
        if m:
            map_path = OUTPUT_DIR / f"Route_{letter}_{safe_name}_map.html"
            m.save(str(map_path))
            print(f"  Map:      {map_path.name}")

        # Manifest
        manifest_path = OUTPUT_DIR / f"Route_{letter}_{safe_name}_manifest.csv"
        write_manifest(ordered, manifest_path)
        print(f"  Manifest: {manifest_path.name}")

        # Kitchen data
        box_counts     = {"Large": 0, "Medium": 0, "Small": 0, "Unknown": 0}
        allergen_notes = []
        for stop in ordered:
            box_counts[stop["box_size"]] = box_counts.get(stop["box_size"], 0) + 1
            if stop["allergens"] and stop["allergens"].lower() not in ("none", ""):
                allergen_notes.append(f"Member {stop['member_id']}: {stop['allergens']}")

        kitchen_rows.append({
            "Route":          f"Route {letter} — {name.replace('_',' ')} ({day})",
            "Total Stops":    len(ordered),
            "Large":          box_counts.get("Large", 0),
            "Medium":         box_counts.get("Medium", 0),
            "Small":          box_counts.get("Small", 0),
            "Unknown":        box_counts.get("Unknown", 0),
            "Allergen Notes": "; ".join(allergen_notes) if allergen_notes else "",
        })

        summary_lines.append(
            f"  Route {letter} ({name.replace('_',' ')}, {day}): "
            f"{len(ordered)} stops | "
            f"L:{box_counts.get('Large',0)} M:{box_counts.get('Medium',0)} "
            f"S:{box_counts.get('Small',0)} | {opt_dist:.1f} mi"
        )

    # Kitchen packing list
    kitchen_path = OUTPUT_DIR / "Kitchen_Packing_List.csv"
    kitchen_headers = ["Route", "Total Stops", "Large", "Medium", "Small", "Unknown", "Allergen Notes"]
    with open(kitchen_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=kitchen_headers)
        writer.writeheader()
        writer.writerows(kitchen_rows)
        if kitchen_rows:
            writer.writerow({
                "Route":        "TOTAL",
                "Total Stops":  sum(r["Total Stops"] for r in kitchen_rows),
                "Large":        sum(r["Large"]       for r in kitchen_rows),
                "Medium":       sum(r["Medium"]      for r in kitchen_rows),
                "Small":        sum(r["Small"]       for r in kitchen_rows),
                "Unknown":      sum(r["Unknown"]     for r in kitchen_rows),
                "Allergen Notes": "",
            })

    # Flags file
    flags_path = OUTPUT_DIR / "Flags.txt"
    with open(flags_path, "w", encoding="utf-8") as f:
        f.write(f"Flags for week of {date.today().isoformat()}\n")
        f.write("=" * 60 + "\n\n")
        if flags_global:
            for line in flags_global:
                f.write(f"  - {line}\n")
        else:
            f.write("  No flags — all clear.\n")

    # Summary
    print(f"\n{'='*60}")
    print(f"  SUMMARY — Week of {date.today().isoformat()}")
    print(f"{'='*60}")
    for line in summary_lines:
        print(line)
    total_stops = sum(r["Total Stops"] for r in kitchen_rows)
    total_L     = sum(r["Large"]       for r in kitchen_rows)
    total_M     = sum(r["Medium"]      for r in kitchen_rows)
    total_S     = sum(r["Small"]       for r in kitchen_rows)
    print(f"\n  TOTAL: {total_stops} stops | Large:{total_L} Medium:{total_M} Small:{total_S}")
    print(f"  Flags:  {len(flags_global)} item(s) to review")
    print(f"\n  Output: {OUTPUT_DIR}")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    main()
