import pandas as pd
import os
import datetime

from common.gebiet_schluessel import gebiet_schluessel
from common.mapping import get_gemeinde_from_gebiet, track_undetected_gebiete, log_missing_gebiete, get_gemeinde_by_schluessel

FILENAME = "geburtsjahrgangsstatistik.xlsx"
INPUT_DIR = "data"
OUTPUT_DIR = "result"
SHEET_NAME = "dadigesamt"

INPUT_FILENAME = os.path.join(INPUT_DIR, FILENAME)
OUTPUT_FILENAME = os.path.join(OUTPUT_DIR, FILENAME)

def parse_excel():
    """
    Main entry point. Processes birth year statistics and produces an Excel summary grouped by Gemeinde and age group.

    Steps:
    - Load Excel file and clean data
    - Map Gebiet (area) to standardized Gemeinde (municipality)
    - Track unmapped Gebiet entries and log them
    - Classify each record into age groups (young, middle, old)
    - Group and sum total population by Gemeinde and age group
    - Generate output Excel file with key demographic indicators
    """
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    if not os.path.exists(INPUT_FILENAME):
        print(f"Error: File {INPUT_FILENAME} not found.")
        return

    xls = pd.ExcelFile(INPUT_FILENAME)
    if SHEET_NAME not in xls.sheet_names:
        print(f"Error: Sheet '{SHEET_NAME}' not found. Available: {xls.sheet_names}")
        return

    # Load data and preprocess
    df = pd.read_excel(xls, sheet_name=SHEET_NAME, dtype=str)
    df = df.dropna(subset=["Gebiet", "Jahrgang"])
    df["Gebiet"] = df["Gebiet"].str.strip()

    # Convert Jahrgang to numeric and drop invalid rows
    df["Jahrgang"] = pd.to_numeric(df["Jahrgang"], errors='coerce')
    df = df.dropna(subset=["Jahrgang"])
    df["Jahrgang"] = df["Jahrgang"].astype(int)

    # Map Gebiet to standardized Gemeinde using the internal mapping logic
    df["Gemeinde"] = df["Gebiet"].map(lambda x: get_gemeinde_from_gebiet(x))

    # Track unmapped Gebiet entries and log warnings
    all_gebieten = df["Gebiet"].dropna().unique().tolist()
    undetected_gebiete = track_undetected_gebiete(all_gebieten)
    log_missing_gebiete(undetected_gebiete)

    if "EW gesamt" not in df.columns:
        print("Error: Column 'EW gesamt' not found.")
        return

    # Convert population to integers and fill missing values with 0
    df["EW gesamt"] = pd.to_numeric(df["EW gesamt"], errors='coerce').fillna(0).astype(int)

    current_year = datetime.datetime.now().year

    def classify_group(year):
        """
        Assigns each person to an age group based on birth year:
        - junge: under 21
        - mittleren: 21 to 64
        - alte: 65 and older
        """
        age = current_year - year
        if age < 21:
            return "junge"
        elif age > 64:
            return "alte"
        return "mittleren"

    df["gruppe"] = df["Jahrgang"].apply(classify_group)

    # Remove entries with missing Gemeinde mapping
    df = df[df["Gemeinde"].notnull() & (df["Gemeinde"] != "")]

    # Aggregate population per Gemeinde and group
    grouped = df.groupby(["Gemeinde", "gruppe"])["EW gesamt"].sum().unstack(fill_value=0).reset_index()

    # Add Gemeinde key and ISO code from lookup table
    grouped["gemeindeschlüssel"] = grouped["Gemeinde"].map(lambda x: gebiet_schluessel.get(x, ("", ""))[0])
    grouped["gemeindeschlüssel"] = grouped["gemeindeschlüssel"].apply(
        lambda x: str(x).zfill(8) if pd.notnull(x) and str(x).isdigit() else ""
    )
    grouped["gemeinde"] = grouped["gemeindeschlüssel"].apply(
        lambda schluessel: get_gemeinde_by_schluessel(schluessel)
    )
    grouped["iso"] = grouped["Gemeinde"].map(lambda x: gebiet_schluessel.get(x, ("", ""))[1])

    # Calculate demographic ratios (young and old per 100 middle-aged)
    grouped["junge quotient"] = (grouped["junge"] / grouped["mittleren"]).replace([float("inf"), -float("inf")], 0) * 100
    grouped["alte quotient"] = (grouped["alte"] / grouped["mittleren"]).replace([float("inf"), -float("inf")], 0) * 100

    # Convert ratios to strings with percent format
    grouped["junge quotient"] = grouped["junge quotient"].round(2).astype(str) + "%"
    grouped["alte quotient"] = grouped["alte quotient"].round(2).astype(str) + "%"

    # Rename columns to more descriptive format
    grouped = grouped.rename(columns={
        "junge": "junge count",
        "alte": "alte count",
        "mittleren": "mittleren count"
    })

    # Remove special aggregate rows
    grouped = grouped[~grouped["gemeinde"].isin(["Ausgewählte Gebiete zusammengefasst", "Sanierungsgebiet"])]

    # Define final column order and save result
    final_columns = [
        "gemeinde",
        "gemeindeschlüssel",
        "iso",
        "junge count",
        "alte count",
        "mittleren count",
        "junge quotient",
        "alte quotient"
    ]

    grouped[final_columns].to_excel(OUTPUT_FILENAME, index=False)
    print(f"Result saved to {OUTPUT_FILENAME}")

if __name__ == "__main__":
    parse_excel()
