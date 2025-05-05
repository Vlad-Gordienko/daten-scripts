import pandas as pd
import os
import time
from geopy.geocoders import Nominatim
from geopy.exc import GeocoderTimedOut

FILENAME = "Altersplanung_Anbieterverzeichnis.xlsx"
INPUT_DIR = "data"
OUTPUT_DIR = "result"
SHEET_NAME = "Anbieter"

INPUT_PATH = os.path.join(INPUT_DIR, FILENAME)
OUTPUT_PATH = os.path.join(OUTPUT_DIR, "anbieter_mit_koordinaten.xlsx")

geolocator = Nominatim(user_agent="superset-mapper")
geocode_cache = {}

def geocode_address(address):
    if address in geocode_cache:
        return geocode_cache[address]
    try:
        location = geolocator.geocode(address)
        if location:
            coords = (location.latitude, location.longitude)
        else:
            coords = (None, None)
    except GeocoderTimedOut:
        time.sleep(1)
        return geocode_address(address)
    geocode_cache[address] = coords
    time.sleep(1)  # пауза между запросами
    return coords

def parse_adressen():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    if not os.path.exists(INPUT_PATH):
        print(f"Error: File {INPUT_PATH} not found.")
        return

    df = pd.read_excel(INPUT_PATH, sheet_name=SHEET_NAME, dtype=str)
    df = df.dropna(subset=["Anschrift", "PLZ", "Ort"])

    df["Anschrift"] = df["Anschrift"].str.strip()
    df["PLZ"] = df["PLZ"].astype(str).str.strip()
    df["Ort"] = df["Ort"].str.strip()

    df["full_address"] = df["Anschrift"] + ", " + df["PLZ"] + " " + df["Ort"]
    df["latitude"] = None
    df["longitude"] = None

    for i, row in df.iterrows():
        coords = geocode_address(row["full_address"])
        df.at[i, "latitude"] = coords[0]
        df.at[i, "longitude"] = coords[1]

    df.to_excel(OUTPUT_PATH, index=False)
    print(f"Result saved to {OUTPUT_PATH}")

if __name__ == "__main__":
    parse_adressen()
