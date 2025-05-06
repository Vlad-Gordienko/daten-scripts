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
OUTPUT_PATH = os.path.join(OUTPUT_DIR, "anbieter_mit_koordinaten.xlsx")

geolocator = Nominatim(user_agent="superset-mapper")
geocode_cache = {}

def geocode_address(address, index, length):
    print(f"{index + 1}/{length}: {address}")
    if address in geocode_cache:
        print('Get from: Cache')
        return geocode_cache[address]
    try:
        print('Get from: Geolocator')
        location = geolocator.geocode(address)
        if location:
            coords = (location.latitude, location.longitude)
        else:
            coords = (None, None)
    except GeocoderTimedOut:
        time.sleep(1)
        return geocode_address(address, index, length)
    geocode_cache[address] = coords
    save_geocode_cache()
    time.sleep(1)
    return coords

def save_geocode_cache():
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
    global geocode_cache

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    if not os.path.exists(INPUT_PATH):
        print(f"Error: File {INPUT_PATH} not found.")
        return

    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, "r", encoding="utf-8") as f:
                geocode_cache = json.load(f)
        except json.JSONDecodeError:
            print("geocode_cache.json corrupted or empty - the cache will be recreated.")
            geocode_cache = {}
    else:
        geocode_cache = {}

    geocode_cache = {str(k): tuple(v) for k, v in geocode_cache.items()}

    df = pd.read_excel(INPUT_PATH, sheet_name=SHEET_NAME, dtype=str)
    df = df.dropna(subset=["Anschrift", "PLZ", "Ort"])
    df = df.drop_duplicates(subset=["Anschrift", "PLZ", "Ort"])
    df.reset_index(drop=True, inplace=True)

    df["Anschrift"] = df["Anschrift"].str.strip()
    df["PLZ"] = df["PLZ"].astype(str).str.strip()
    df["Ort"] = df["Ort"].str.strip()

    df["full_address"] = df["Anschrift"] + ", " + df["PLZ"] + " " + df["Ort"]
    df["latitude"] = None
    df["longitude"] = None

    length = len(df)
    for i, row in df.iterrows():
        coords = geocode_address(row["full_address"], i, length)
        df.at[i, "latitude"] = coords[0]
        df.at[i, "longitude"] = coords[1]

    columns_to_keep = [
        "Anbieter",
        "Anschrift",
        "PLZ",
        "Ort",
        "Telefon",
        "E-Mail",
        "full_address",
        "latitude",
        "longitude"
    ]
    df = df[columns_to_keep]

    df.to_excel(OUTPUT_PATH, index=False)
    print(f"Result saved to {OUTPUT_PATH}")

if __name__ == "__main__":
    parse_adressen()
