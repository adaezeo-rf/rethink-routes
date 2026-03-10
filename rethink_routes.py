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
import os
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

# Route definitions: each entry is (display_name, borough_hint, [zip codes])
# borough_hint is used to build geocoding query fallbacks
ROUTES = [
    ("A", "Jamaica_SE_Queens",    "Queens",    ["11412","11417","11418","11419","11420","11423","11429",
                                                "11432","11433","11434","11435","11436","11453"]),
    ("B", "Far_Rockaway",         "Queens",    ["11691","11692"]),
    ("C", "Central_NW_Queens",    "Queens",    ["11355","11356","11365","11368","11369","11370",
                                                "11372","11377","11379","11385"]),
    ("D", "Bronx_South",          "Bronx",     ["10451","10452","10453","10454","10455","10456","10459","10460"]),
    ("E", "Bronx_North",          "Bronx",     ["10461","10462","10463","10466","10467","10468","10469","10472","10473","10475"]),
    ("F", "Upper_Manhattan",      "Manhattan", ["10024","10025","10027","10029","10030","10031","10032",
                                                "10033","10034","10039","10040","10129"]),
    ("G", "Midtown_Manhattan",    "Manhattan", ["10001","10002","10009","10011","10023","10036","10044","10075"]),
    ("H", "Brooklyn",             "Brooklyn",  ["11101","11103","11106","11201","11203","11206","11207","11208","11209",
                                                "11210","11211","11212","11214","11216","11217","11219","11220","11221",
                                                "11223","11224","11226","11228","11231","11233","11235","11236","11237",
                                                "11238","11239","11249"]),
]

BOX_COLORS = {
    "Large":   "#e74c3c",
    "Medium":  "#e67e22",
    "Small":   "#2980b9",
    "Unknown": "#7f8c8d",
}

# ── Helpers ───────────────────────────────────────────────────────────────────

def clean_box(raw):
    if not raw:
        return "Unknown"
    s = str(raw).strip().lower()
    if s.startswith("l"): return "Large"
    if s.startswith("m"): return "Medium"
    if s.startswith("s"): return "Small"
    return str(raw).strip().title()


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


def nearest_neighbor(stops):
    n = len(stops)
    unvisited = set(range(1, n))
    route = [0]
    while unvisited:
        last = route[-1]
        nearest = min(unvisited, key=lambda j: haversine_miles(stops[last]["latlon"], stops[j]["latlon"]))
        route.append(nearest)
        unvisited.remove(nearest)
    return route


def two_opt(route, stops):
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


def optimize_route(stops):
    """Return stops reordered by nearest-neighbor + 2-opt. Returns original order if < 2 stops."""
    if len(stops) < 2:
        return stops, 0.0, 0.0
    n = len(stops)
    orig_dist = route_distance(list(range(n)), stops)
    nn = nearest_neighbor(stops)
    opt = two_opt(nn, stops)
    opt_dist = route_distance(opt, stops)
    return [stops[i] for i in opt], orig_dist, opt_dist


# ── Geocoding ─────────────────────────────────────────────────────────────────

def load_cache():
    if CACHE_FILE.exists():
        with open(CACHE_FILE, "r") as f:
            raw = json.load(f)
        # JSON keys are strings; convert back to (addr, zip) tuples
        return {tuple(k.split("|||")): tuple(v) if v else None for k, v in raw.items()}
    return {}


def save_cache(cache):
    raw = {"|||".join(k): list(v) if v else None for k, v in cache.items()}
    with open(CACHE_FILE, "w") as f:
        json.dump(raw, f, indent=2)


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

def build_map(ordered_stops, route_letter, route_name, opt_dist):
    if not ordered_stops:
        return None

    lats = [s["latlon"][0] for s in ordered_stops]
    lons = [s["latlon"][1] for s in ordered_stops]
    center = [sum(lats) / len(lats), sum(lons) / len(lons)]

    m = folium.Map(location=center, zoom_start=13, tiles="OpenStreetMap")

    # Title
    title = (
        f'<div style="position:fixed;top:12px;left:50%;transform:translateX(-50%);'
        f'z-index:1000;background:white;padding:8px 20px;border-radius:6px;'
        f'border:1px solid #ccc;box-shadow:0 2px 8px rgba(0,0,0,.3);'
        f'font-family:Arial,sans-serif;font-size:15px;font-weight:bold;">'
        f'Route {route_letter} &mdash; {route_name.replace("_", " ")} '
        f'({len(ordered_stops)} stops &bull; {opt_dist:.1f} mi)</div>'
    )
    m.get_root().html.add_child(folium.Element(title))

    # Polyline
    coords = [s["latlon"] for s in ordered_stops]
    folium.PolyLine(
        locations=coords, color="#2c3e50", weight=2.5, opacity=0.75,
        tooltip=f"Route {route_letter}"
    ).add_to(m)

    # Markers
    for i, stop in enumerate(ordered_stops, start=1):
        color = BOX_COLORS.get(stop["box_size"], "#7f8c8d")
        lat, lon = stop["latlon"]

        # Flag anomalies visually
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
        flag_html = f'<br><b style="color:#e74c3c;">⚠ FLAG: {stop["flag"]}</b>' if stop["flag"] else ""

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
        'Flagged stop</div>'
    )
    m.get_root().html.add_child(folium.Element(legend_html))

    return m


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

    # Load Excel
    print(f"Loading: {EXCEL_PATH}")
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
    all_stops = []
    flags_global = []

    for row in rows[1:]:
        status = str(row[col["Status"]] or "").strip().lower()
        if status != "active":
            continue

        zipcode = str(row[col["Zip"]] or "").replace(".0", "").strip().zfill(5)
        addr1   = str(row[col["Address Line 1"]] or "").strip()
        addr2   = str(row[col["Address Line 2"]] or "").strip()
        city    = str(row[col["City"]]   or "").strip()
        state   = str(row[col["State"]]  or "").strip()

        display_addr = addr1
        if addr2 and addr2.lower() not in ("none", ""):
            display_addr += f", {addr2}"
        display_addr += f", {city}, {state} {zipcode}"

        allergens = str(row[col.get("Meal Preferences/Allergens", -1)] or "").strip() \
            if col.get("Meal Preferences/Allergens") is not None else ""
        avail     = str(row[col.get("Available Delivery Days", -1)] or "").strip() \
            if col.get("Available Delivery Days") is not None else ""
        notes_raw = " | ".join(filter(None, [
            str(row[col[c]] or "").strip()
            for c in ["Unnamed: 13", "Unnamed: 14", "Unnamed: 15"]
            if c in col
        ]))

        # Auto-flag anomalies
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
    print(f"Pre-flight flags:     {len(flags_global)}\n")

    # Assign stops to routes
    zip_to_route = {}
    for letter, name, borough, zips in ROUTES:
        for z in zips:
            zip_to_route[z] = (letter, name, borough)

    unassigned = []
    route_stops = {letter: [] for letter, *_ in ROUTES}

    for stop in all_stops:
        assignment = zip_to_route.get(stop["zipcode"])
        if assignment:
            route_stops[assignment[0]].append((stop, assignment[2]))  # (stop, borough)
        else:
            unassigned.append(stop)
            flags_global.append(
                f"Member {stop['member_id']} (zip {stop['zipcode']}) not assigned to any route"
            )

    if unassigned:
        print(f"WARNING: {len(unassigned)} member(s) not assigned to a route (new zip codes?)")
        for s in unassigned:
            print(f"  - {s['member_id']} zip={s['zipcode']} {s['addr1']}")

    # Load geocode cache
    geocache = load_cache()
    cached_before = len(geocache)
    geolocator = Nominatim(user_agent="rethink_food_router_v2")

    # Output folder
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    print(f"Output folder: {OUTPUT_DIR}\n")

    # Per-route processing
    kitchen_rows = []
    summary_lines = []

    for letter, name, borough, zips in ROUTES:
        entries = route_stops[letter]
        if not entries:
            print(f"Route {letter} ({name}): no active members, skipping.")
            continue

        stops_for_route = [e[0] for e in entries]
        borough_hint    = entries[0][1]

        print(f"{'-'*50}")
        print(f"Route {letter} — {name.replace('_',' ')} ({len(stops_for_route)} members)")

        # Geocode
        geocoded = []
        for stop in stops_for_route:
            key = (stop["addr1"], stop["zipcode"])
            if key in geocache:
                stop["latlon"] = geocache[key]
                if stop["latlon"]:
                    geocoded.append(stop)
                    print(f"  [cached] {stop['addr1']}")
                else:
                    print(f"  [cached-FAIL] {stop['addr1']}")
            else:
                print(f"  [geocoding] {stop['addr1']} ({stop['zipcode']})...")
                result = geocode_stop(geolocator, stop["addr1"], stop["zipcode"], borough_hint, geocache)
                stop["latlon"] = result
                if result:
                    geocoded.append(stop)
                    print(f"    -> {result}")
                else:
                    print(f"    -> FAILED (will be excluded from map)")
                    flags_global.append(
                        f"Member {stop['member_id']} ({stop['addr1']}, {stop['zipcode']}): geocoding failed"
                    )
                time.sleep(1)

        if not geocoded:
            print(f"  No geocoded stops — skipping route {letter}.")
            continue

        # Optimize
        ordered, orig_dist, opt_dist = optimize_route(geocoded)
        saved = orig_dist - opt_dist
        pct   = (saved / orig_dist * 100) if orig_dist > 0 else 0
        print(f"  Distance: {orig_dist:.1f} mi -> {opt_dist:.1f} mi (saved {saved:.1f} mi / {pct:.0f}%)")

        # Map
        safe_name = name.replace(" ", "_")
        map_path  = OUTPUT_DIR / f"Route_{letter}_{safe_name}_map.html"
        m = build_map(ordered, letter, name, opt_dist)
        if m:
            m.save(str(map_path))
            print(f"  Map:      {map_path.name}")

        # Manifest
        manifest_path = OUTPUT_DIR / f"Route_{letter}_{safe_name}_manifest.csv"
        write_manifest(ordered, manifest_path)
        print(f"  Manifest: {manifest_path.name}")

        # Kitchen packing data
        box_counts = {"Large": 0, "Medium": 0, "Small": 0, "Unknown": 0}
        allergen_notes = []
        for stop in ordered:
            box_counts[stop["box_size"]] = box_counts.get(stop["box_size"], 0) + 1
            if stop["allergens"] and stop["allergens"].lower() not in ("none", ""):
                allergen_notes.append(f"Member {stop['member_id']}: {stop['allergens']}")

        kitchen_rows.append({
            "Route":         f"Route {letter} — {name.replace('_',' ')}",
            "Total Stops":   len(ordered),
            "Large":         box_counts.get("Large", 0),
            "Medium":        box_counts.get("Medium", 0),
            "Small":         box_counts.get("Small", 0),
            "Unknown":       box_counts.get("Unknown", 0),
            "Allergen Notes": "; ".join(allergen_notes) if allergen_notes else "",
        })

        summary_lines.append(
            f"  Route {letter} ({name.replace('_',' ')}): "
            f"{len(ordered)} stops | "
            f"L:{box_counts.get('Large',0)} M:{box_counts.get('Medium',0)} S:{box_counts.get('Small',0)} | "
            f"{opt_dist:.1f} mi optimized"
        )

    # Kitchen packing list
    kitchen_path = OUTPUT_DIR / "Kitchen_Packing_List.csv"
    kitchen_headers = ["Route", "Total Stops", "Large", "Medium", "Small", "Unknown", "Allergen Notes"]
    with open(kitchen_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=kitchen_headers)
        writer.writeheader()
        writer.writerows(kitchen_rows)

        # Totals row
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
    print(f"\n{'─'*50}")
    print(f"Kitchen packing list: {kitchen_path.name}")

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
    print(f"Flags file:           {flags_path.name}")

    # Summary
    new_in_cache = len(geocache) - cached_before
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
    print(f"  Geocache: {new_in_cache} new address(es) added ({len(geocache)} total cached)")
    print(f"  Flags:    {len(flags_global)} item(s) to review")
    print(f"\n  Output: {OUTPUT_DIR}")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    main()
