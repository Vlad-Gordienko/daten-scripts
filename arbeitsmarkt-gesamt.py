import pandas as pd
import os
import re
import logging
from typing import List

from common.mapping import normalize_gemeinde_name

INPUT_DIR = "data/arbeitsortbeschäftigung"
GEMBAND_DIR = "data/gemband"
OUTPUT_DIR = "result"
OUTPUT_FILENAME = "arbeitsmarkt_gesamt.xlsx"

YEARS = [2020, 2021, 2022, 2023, 2024]

GEMEINDE_ID_MAP = {
    "Altenstadt": 1,
    "Bad Nauheim": 2,
    "Bad Vilbel": 3,
    "Büdingen": 4,
    "Butzbach": 5,
    "Echzell": 6,
    "Florstadt": 7,
    "Friedberg (Hessen)": 8,
    "Gedern": 9,
    "Glauburg": 10,
    "Hirzenhain": 11,
    "Karben": 12,
    "Kefenrod": 13,
    "Limeshain": 14,
    "Münzenberg": 15,
    "Nidda": 16,
    "Niddatal": 17,
    "Ober-Mörlen": 18,
    "Ortenberg": 19,
    "Ranstadt": 20,
    "Reichelsheim (Wetterau)": 21,
    "Rockenberg": 22,
    "Rosbach v. d. Höhe": 23,
    "Wölfersheim": 24,
    "Wöllstadt": 25,
    "Wetteraukreis": 0
}

# Mapping of unemployment and sector row indexes (constant across years)
ROW_INDEXES_ARBEITSLOSIGKEIT = {
    "total": 37,                 # Insgesamt
    "male-arbeitslos": 38,       # Männer (Arbeitslos)
    "female-arbeitslos": 39,     # Frauen (Arbeitslos)
    "sgb3": 44,                  # SGB III
    "sgb2": 45,                  # SGB II
}

ROW_INDEXES_SEKTOREN = {
    "nace-a": 17,    # Land- und Forstwirtschaft, Fischerei ( A )
    "nace-b-f": 18,  # Produzierendes Gewerbe ( B - F )
    "nace-g-i": 19,  # Handel, Verkehr und Gastgewerbe ( G - I )
    "nace-j-u": 20,  # Sonstige Dienstleistungen ( J - U )
}

GEMBAND_COL_MAP = {
    "male-pendler": 3,   # D -> Männer (Pendlersaldo)
    "female-pendler": 4, # E -> Frauen (Pendlersaldo)
    "einpendler": 12,      # M -> Einpendler
    "auspendler": 13,     # N -> Auspendler
}

def parse_value(v):
    # Safely convert value to float; fallback to 0 on failure
    try:
        return float(v)
    except Exception:
        return 0.0

def extract_arbeitsmarkt_data(file_path: str) -> pd.DataFrame:
    # Parse municipality name from filename
    # Read values for each year and append to result
    df = pd.read_excel(file_path, sheet_name="Daten", header=None)

    filename = os.path.basename(file_path)
    match = re.match(r"Arbeitsmarkt-kommunal_\d+_(.*)\.xlsx", filename)
    gemeinde_raw = match.group(1).replace("_", " ") if match else "Unbekannt"
    gemeinde = normalize_gemeinde_name(gemeinde_raw)

    rows = []
    for year_idx, year in enumerate(YEARS):
        row = {
            "gemeinde": gemeinde,
            "gemeinde_id": int(GEMEINDE_ID_MAP.get(gemeinde, 0)),
            "year": year
        }

        for colname, row_index in ROW_INDEXES_ARBEITSLOSIGKEIT.items():
            value = parse_value(df.iloc[row_index, 2 + year_idx])
            row[colname] = int(round(value))

        for colname, row_index in ROW_INDEXES_SEKTOREN.items():
            value = parse_value(df.iloc[row_index, 2 + year_idx])
            row[colname] = int(round(value))

        rows.append(row)

    return pd.DataFrame(rows)


def extract_gemband_data(file_path: str) -> pd.DataFrame:
    # Determine engine for .xlsb support
    # Extract year from filename (used to select row range)

    # Define year-specific row ranges due to inconsistent file structures
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
        2020: (2947, 2971), # (2948–2972)
        2021: (2947, 2971), # (2948–2972)
        2022: (2945, 2969), # (2946–2970)
        2023: (2942, 2966), # (2943–2967)
        2024: (2942, 2966), # (2943–2967)
    }

    # Iterate over rows and normalize municipality names
    # Skip invalid or unknown entries

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
        if not gemeinde or gemeinde == "Unbekannt":
            continue

        entry = {
            "gemeinde": gemeinde,
            "gemeinde_id": int(GEMEINDE_ID_MAP.get(gemeinde, 0)),
            "year": year
        }
        for name, col_idx in GEMBAND_COL_MAP.items():
            entry[name] = parse_value(row[col_idx])

        result.append(entry)

    return pd.DataFrame(result)


def main():
    # Load and process all arbeitsmarkt files
    # Load and process all gemband files

    # Merge datasets on Gemeinde and Jahr if available
    # Compute yearly sum row for the entire district (Wetteraukreis)
    # Save final result to Excel
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    arbeitsmarkt_frames: List[pd.DataFrame] = []
    for file in sorted(os.listdir(INPUT_DIR)):
        if not (file.endswith(".xlsx") and file.startswith("Arbeitsmarkt-kommunal")):
            continue
        path = os.path.join(INPUT_DIR, file)
        try:
            arbeitsmarkt_frames.append(extract_arbeitsmarkt_data(path))
        except Exception as e:
            print(f"Fehler bei {file}: {e}")

    if not arbeitsmarkt_frames:
        print("No 'arbeitsmarkt' data")
        return

    df_arbeitsmarkt = pd.concat(arbeitsmarkt_frames, ignore_index=True)

    gemband_frames: List[pd.DataFrame] = []
    for file in sorted(os.listdir(GEMBAND_DIR)):
        if not file.endswith((".xlsx", ".xlsb")):
            continue
        path = os.path.join(GEMBAND_DIR, file)
        try:
            gemband_frames.append(extract_gemband_data(path))
        except Exception as e:
            print(f"Error. Read file {file}: {e}")

    if gemband_frames:
        df_gemband = pd.concat(gemband_frames, ignore_index=True)
        if df_gemband.empty:
            print("No gemband data")
            return
        df_merged = pd.merge(
            df_arbeitsmarkt,
            df_gemband,
            how="left",
            on=["gemeinde", "gemeinde_id", "year"]
        )
    else:
        print("No valid gemband data extracted.")
        df_merged = df_arbeitsmarkt

    #     sum_df = df_merged.groupby("year").sum(numeric_only=True).reset_index()
    #     sum_df["gemeinde"] = "Wetteraukreis"
    #
    #     metric_cols = [c for c in df_merged.columns if c not in ["gemeinde", "year"]]
    #     sum_df = sum_df[["gemeinde", "year"] + metric_cols]
    #
    #     final_df = pd.concat([df_merged, sum_df], ignore_index=True)

    output_path = os.path.join(OUTPUT_DIR, OUTPUT_FILENAME)
    df_merged.to_excel(output_path, index=False)
    print(f"Result saved to: {output_path}")


if __name__ == "__main__":
    main()
