import pandas as pd
import os
import time
import json
from geopy.geocoders import Nominatim
from geopy.exc import GeocoderTimedOut

FILENAME = "Altersplanung_Anbieterverzeichnis.xlsx"
INPUT_DIR = "data"
OUTPUT_DIR = "result"
SHEET_NAME = "Anbieter"
CACHE_FILE = "geocode_cache.json"

INPUT_PATH = os.path.join(INPUT_DIR, FILENAME)
OUTPUT_PATH = os.path.join(OUTPUT_DIR, "altersplanung.xlsx")

# Initialize geocoder with custom user-agent to avoid being blocked
geolocator = Nominatim(user_agent="superset-mapper")

# Dictionary to store geocoded address -> (latitude, longitude)
geocode_cache = {}

def geocode_address(address, index, total):
    """
    Attempts to geocode a full address.
    Uses a local cache to avoid redundant requests.
    Implements retry logic for timeouts and respects rate-limiting (sleep).
    """
    print(f"{index + 1}/{total}: {address}")

    if address in geocode_cache:
        print("---> Get from: Cache")
        return geocode_cache[address]

    print("---> Get from: *Geolocator")
    try:
        location = geolocator.geocode(address)
        if location:
            coords = (location.latitude, location.longitude)
            geocode_cache[address] = coords
            save_geocode_cache()  # persist new entry
            time.sleep(1)  # avoid hitting rate limits
            return coords
        else:
            print(f"(!) Address not found: {address} (!)")
            print('-' * 28)
            return (None, None)
    except GeocoderTimedOut:
        # Retry once if timeout occurs
        print("Timeout occurred, retrying...")
        time.sleep(1)
        return geocode_address(address, index, total)


def save_geocode_cache():
    """
    Reads existing geocode cache from file, updates it with new entries,
    and writes back to disk. Ensures cache is preserved between runs.
    """
    existing_cache = {}

    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, "r", encoding="utf-8") as f:
                existing_cache = json.load(f)
        except json.JSONDecodeError:
            print("geocode_cache.json corrupted or empty - the cache will be recreated.")

    updated_cache = {**existing_cache, **geocode_cache}

    with open(CACHE_FILE, "w", encoding="utf-8") as f:
        json.dump(updated_cache, f, ensure_ascii=False, indent=2)

def parse_adressen():
    """
    Main parsing logic:
    - Loads address data from Excel
    - Constructs full addresses
    - Geocodes each address with caching
    - Outputs a new Excel with lat/lng appended
    """
    global geocode_cache

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    if not os.path.exists(INPUT_PATH):
        print(f"Error: File {INPUT_PATH} not found.")
        return

    # Load existing geocode cache if present
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, "r", encoding="utf-8") as f:
                geocode_cache = json.load(f)
        except json.JSONDecodeError:
            print("geocode_cache.json corrupted or empty - the cache will be recreated.")
            geocode_cache = {}
    else:
        geocode_cache = {}

    # Convert all coordinates from list to tuple format
    geocode_cache = {str(k): tuple(v) for k, v in geocode_cache.items()}

    # Read data
    df = pd.read_excel(INPUT_PATH, sheet_name=SHEET_NAME, dtype=str)

    # Ensure required fields are present and cleaned
    df = df.dropna(subset=["Anschrift", "PLZ", "Ort"])
    df = df.drop_duplicates(subset=["Anschrift", "PLZ", "Ort"])
    df.reset_index(drop=True, inplace=True)

    # Normalize address columns
    df["Anschrift"] = df["Anschrift"].str.strip()
    df["PLZ"] = df["PLZ"].astype(str).str.strip()
    df["Ort"] = df["Ort"].str.strip()

    # Build full address for geocoding
    df["Address"] = df["Anschrift"] + ", " + df["PLZ"] + " " + df["Ort"]
    df["latitude"] = None
    df["longitude"] = None

    # Geocode each address row by row
    length = len(df)
    for i, row in df.iterrows():
        coords = geocode_address(row["Address"], i, length)
        df.at[i, "latitude"] = coords[0]
        df.at[i, "longitude"] = coords[1]

    # Keep only relevant columns
    columns_to_keep = [
        "Branche",
        "Anbieter",
        "Anschrift",
        "Address",
        "latitude",
        "longitude"
    ]
    df = df[columns_to_keep]

    df.to_excel(OUTPUT_PATH, index=False)
    print(f"Result saved to {OUTPUT_PATH}")

# Entrypoint for CLI execution
if __name__ == "__main__":
    parse_adressen()
