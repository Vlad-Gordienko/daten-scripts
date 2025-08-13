import pandas as pd
import os
import re
import logging
from typing import List
from common.mapping import normalize_gemeinde_name

INPUT_DIR = "data/arbeitsortbeschäftigung"
GEMBAND_DIR = "data/gemband"
OUTPUT_DIR = "result"
OUTPUT_FILENAME = "arbeitsmarkt_gesamt_2.xlsx"

YEARS = [2020, 2021, 2022, 2023, 2024]

ROW_INDEXES_ARBEITSLOSIGKEIT = {
    "Insgesamt": 37,
    "Männer": 38,
    "Frauen": 39,
    "SGB III": 44,
    "SGB II": 45,
}
ROW_INDEXES_SEKTOREN = {
    "Land- und Forstwirtschaft, Fischerei ( A )": 17,
    "Produzierendes Gewerbe ( B - F )": 18,
    "Handel, Verkehr und Gastgewerbe ( G - I )": 19,
    "Sonstige Dienstleistungen ( J - U )": 20,
}

def parse_value(v):
    try:
        return float(v)
    except:
        return 0

def extract_arbeitsmarkt_data(file_path: str) -> pd.DataFrame:
    df = pd.read_excel(file_path, sheet_name="Daten", header=None)
    filename = os.path.basename(file_path)
    match = re.match(r"Arbeitsmarkt-kommunal_\d+_(.*)\.xlsx", filename)
    gemeinde_raw = match.group(1).replace("_", " ") if match else "Unbekannt"
    gemeinde = normalize_gemeinde_name(gemeinde_raw)
    rows = []

    for year_idx, year in enumerate(YEARS):
        row = {"Gemeinde": gemeinde, "Jahr": year}
        for colname, row_index in ROW_INDEXES_ARBEITSLOSIGKEIT.items():
            row[colname] = int(round(parse_value(df.iloc[row_index, 2 + year_idx])))
        for colname, row_index in ROW_INDEXES_SEKTOREN.items():
            row[colname] = int(round(parse_value(df.iloc[row_index, 2 + year_idx])))
        rows.append(row)

    return pd.DataFrame(rows)

def extract_gemband_data(file_path: str) -> pd.DataFrame:
    import warnings
    warnings.simplefilter("ignore")
    _, ext = os.path.splitext(file_path)
    engine = "pyxlsb" if ext == ".xlsb" else None
    df = pd.read_excel(file_path, sheet_name="Gemeindedaten", engine=engine, header=None)

    filename = os.path.basename(file_path)
    match = re.search(r"0-(\d{4})06", filename)
    year = int(match.group(1)) if match else None
    if not year:
        raise ValueError(f"Cannot extract year from filename: {filename}")

    row_range_by_year = {
        2020: (2947, 2971),
        2021: (2947, 2971),
        2022: (2945, 2969),
        2023: (2942, 2966),
        2024: (2942, 2966),
    }
    if year not in row_range_by_year:
        raise ValueError(f"No row mapping for year: {year}")

    row_start, row_end = row_range_by_year[year]
    result = []
    for idx in range(row_start, row_end + 1):
        row = df.iloc[idx]
        gemeinde_raw = str(row[1]).strip()
        if not gemeinde_raw or gemeinde_raw.lower() == "nan":
            continue
        gemeinde = normalize_gemeinde_name(gemeinde_raw)
        if gemeinde is None or gemeinde == "Unbekannt":
            continue

        entry = {
            "Gemeinde": gemeinde,
            "Jahr": year,
            "Männer (Pendlersaldo)": parse_value(row[3]),
            "Frauen (Pendlersaldo)": parse_value(row[4]),
            "Einpendler": parse_value(row[12]),
            "Auspendler": parse_value(row[13]),
        }
        result.append(entry)

    return pd.DataFrame(result)

def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    arbeitsmarkt_frames: List[pd.DataFrame] = []
    for file in sorted(os.listdir(INPUT_DIR)):
        path = os.path.join(INPUT_DIR, file)
        if file.endswith(".xlsx") and file.startswith("Arbeitsmarkt-kommunal"):
            try:
                arbeitsmarkt_frames.append(extract_arbeitsmarkt_data(path))
            except Exception as e:
                print(f"Fehler bei {file}: {e}")

    gemband_frames: List[pd.DataFrame] = []
    for file in sorted(os.listdir(GEMBAND_DIR)):
        path = os.path.join(GEMBAND_DIR, file)
        if file.endswith((".xlsx", ".xlsb")):
            try:
                gemband_frames.append(extract_gemband_data(path))
            except Exception as e:
                print(f"Error. Read file {file}: {e}")

    if not arbeitsmarkt_frames:
        print("No 'arbeitsmarkt' data")
        return

    df_arbeitsmarkt = pd.concat(arbeitsmarkt_frames, ignore_index=True)

    if gemband_frames:
        df_gemband = pd.concat(gemband_frames, ignore_index=True)
        if df_gemband.empty:
            print("No data")
            return
        df_merged = pd.merge(df_arbeitsmarkt, df_gemband, how="left", on=["Gemeinde", "Jahr"])
    else:
        df_merged = df_arbeitsmarkt

    sum_df = df_merged.groupby("Jahr").sum(numeric_only=True).reset_index()
    sum_df["Gemeinde"] = "Wetteraukreis"
    columns = ["Gemeinde", "Jahr"] + [col for col in df_merged.columns if col not in ["Gemeinde", "Jahr"]]
    sum_df = sum_df[columns]
    final_df = pd.concat([df_merged, sum_df], ignore_index=True)

    long_df = final_df.melt(id_vars=["Gemeinde", "Jahr"],
                            var_name="Eigenschaft",
                            value_name="Anzahl")

    output_path = os.path.join(OUTPUT_DIR, OUTPUT_FILENAME)
    long_df.to_excel(output_path, index=False)
    print(f"Result saved to: {output_path}")

if __name__ == "__main__":
    main()
