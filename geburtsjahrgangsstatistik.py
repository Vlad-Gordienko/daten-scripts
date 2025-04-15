import pandas as pd
import os
import datetime

from common.gebiet_schluessel import gebiet_schluessel
from common.mapping import get_gemeinde_from_gebiet, track_undetected_gebiete, log_missing_gebiete

FILENAME = "geburtsjahrgangsstatistik.xlsx"
INPUT_DIR = "data"
OUTPUT_DIR = "result"
SHEET_NAME = "dadigesamt"

INPUT_FILENAME = os.path.join(INPUT_DIR, FILENAME)
OUTPUT_FILENAME = os.path.join(OUTPUT_DIR, FILENAME)

def parse_excel():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    if not os.path.exists(INPUT_FILENAME):
        print(f"Error: File {INPUT_FILENAME} not found.")
        return

    xls = pd.ExcelFile(INPUT_FILENAME)
    if SHEET_NAME not in xls.sheet_names:
        print(f"Error: Sheet '{SHEET_NAME}' not found. Available: {xls.sheet_names}")
        return

    df = pd.read_excel(xls, sheet_name=SHEET_NAME, dtype=str)
    df = df.dropna(subset=["Gebiet", "Jahrgang"])
    df["Gebiet"] = df["Gebiet"].str.strip()

    df["Jahrgang"] = pd.to_numeric(df["Jahrgang"], errors='coerce')
    df = df.dropna(subset=["Jahrgang"])
    df["Jahrgang"] = df["Jahrgang"].astype(int)

    df["Gemeinde"] = df["Gebiet"].map(lambda x: get_gemeinde_from_gebiet(x))

    all_gebieten = df["Gebiet"].dropna().unique().tolist()
    undetected_gebiete = track_undetected_gebiete(all_gebieten)
    log_missing_gebiete(undetected_gebiete)

    if "EW gesamt" not in df.columns:
        print("Error: Column 'EW gesamt' not found.")
        return

    df["EW gesamt"] = pd.to_numeric(df["EW gesamt"], errors='coerce').fillna(0).astype(int)

    current_year = datetime.datetime.now().year

    def classify_group(year):
        age = current_year - year
        if age < 21:
            return "junge"
        elif age > 64:
            return "alte"
        return "mittleren"

    df["gruppe"] = df["Jahrgang"].apply(classify_group)

    grouped = df.groupby(["Gemeinde", "gruppe"])["EW gesamt"].sum().unstack(fill_value=0).reset_index()
    grouped["amtlicher gemeindeschlüssel"] = grouped["Gemeinde"].map(lambda x: gebiet_schluessel.get(x, ("", ""))[0])
    grouped["iso"] = grouped["Gemeinde"].map(lambda x: gebiet_schluessel.get(x, ("", ""))[1])

    grouped["junge quotient"] = (grouped["junge"] / grouped["mittleren"]).replace([float("inf"), -float("inf")], 0) * 100
    grouped["alte quotient"] = (grouped["alte"] / grouped["mittleren"]).replace([float("inf"), -float("inf")], 0) * 100

    grouped["junge quotient"] = grouped["junge quotient"].round(2).astype(str) + "%"
    grouped["alte quotient"] = grouped["alte quotient"].round(2).astype(str) + "%"

    grouped = grouped.rename(columns={
        "Gemeinde": "gemeinde",
        "junge": "junge count",
        "alte": "alte count",
        "mittleren": "mittleren count"
    })

    grouped = grouped[~grouped["gemeinde"].isin(["Ausgewählte Gebiete zusammengefasst", "Sanierungsgebiet"])]

    final_columns = [
        "gemeinde",
        "amtlicher gemeindeschlüssel",
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
