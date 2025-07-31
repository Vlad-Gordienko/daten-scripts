import pandas as pd
import os
import re
from typing import List

from common.mapping import normalize_gemeinde_name

INPUT_DIR = "data/arbeitsortbeschäftigung"
OUTPUT_DIR = "result"
OUTPUT_FILENAME = "arbeitsmarkt_gesamt.xlsx"

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


def extract_data(file_path: str) -> pd.DataFrame:
    df = pd.read_excel(file_path, sheet_name="Daten", header=None)

    filename = os.path.basename(file_path)
    match = re.match(r"Arbeitsmarkt-kommunal_\d+_(.*)\.xlsx", filename)
    gemeinde_raw = match.group(1).replace("_", " ") if match else "Unbekannt"
    gemeinde = normalize_gemeinde_name(gemeinde_raw)

    rows = []

    for year_idx, year in enumerate(YEARS):
        row = {"Gemeinde": gemeinde, "Jahr": year}

        for colname, row_index in ROW_INDEXES_ARBEITSLOSIGKEIT.items():
            value = parse_value(df.iloc[row_index, 2 + year_idx])
            row[colname] = int(round(value))

        for colname, row_index in ROW_INDEXES_SEKTOREN.items():
            value = parse_value(df.iloc[row_index, 2 + year_idx])
            row[colname] = int(round(value))

        rows.append(row)

    return pd.DataFrame(rows)


def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    all_files = [
        os.path.join(INPUT_DIR, f)
        for f in os.listdir(INPUT_DIR)
        if f.endswith(".xlsx") and f.startswith("Arbeitsmarkt-kommunal")
    ]

    result_frames: List[pd.DataFrame] = []
    for file in sorted(all_files):
        try:
            df = extract_data(file)
            result_frames.append(df)
        except Exception as e:
            print(f"Ошибка при обработке {file}: {e}")

    if not result_frames:
        print("Нет данных для обработки.")
        return

    final_df = pd.concat(result_frames, ignore_index=True)

    sum_df = final_df.groupby("Jahr").sum(numeric_only=True).reset_index()
    sum_df["Gemeinde"] = "Wetteraukreis"

    columns = ["Gemeinde", "Jahr"] + [col for col in final_df.columns if col not in ["Gemeinde", "Jahr"]]
    sum_df = sum_df[columns]
    final_df = pd.concat([final_df, sum_df], ignore_index=True)

    output_path = os.path.join(OUTPUT_DIR, OUTPUT_FILENAME)
    final_df.to_excel(output_path, index=False)
    print(f"Result saved to : {output_path}")


if __name__ == "__main__":
    main()
