import pandas as pd
import os
import re

INPUT_DIR = "data/arbeitsortbeschäftigung"
OUTPUT_DIR = "result"
OUTPUT_FILENAME = "arbeitslose_wetterau.xlsx"

RENAME_MAP = {
    "Arbeitslose im Rechtskreis SGB III": "SGB III",
    "Arbeitslose im Rechtskreis SGB II": "SGB II"
}

def extract_fixed(file_path: str) -> pd.DataFrame:
    df = pd.read_excel(file_path, sheet_name="Daten", header=None)

    years = [2020, 2021, 2022, 2023, 2024]

    categories_raw = df.iloc[38:46, 1].tolist()
    categories = ["Insgesamt"] + [str(c).strip() for c in categories_raw]

    values = df.iloc[37:46, 2:7].values.tolist()
    data_matrix = [pd.Series(row).fillna(0).tolist() for row in values]

    n_years = len(years)
    for idx in range(len(data_matrix)):
        data_matrix[idx] = data_matrix[idx][:n_years]

    filename = os.path.basename(file_path)
    match = re.match(r"Arbeitsmarkt-kommunal_\d+_(.*)\.xlsx", filename)
    gemeinde = match.group(1).replace("_", " ") if match else "Unbekannt"

    rows = []
    for col_idx, year in enumerate(years):
        row = {"Gemeinde": gemeinde, "Jahr": year}
        for cat_idx, category in enumerate(categories):
            cat_clean = RENAME_MAP.get(category, category)
            val = data_matrix[cat_idx][col_idx]
            row[cat_clean] = int(round(val))
        rows.append(row)

    return pd.DataFrame(rows)

def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    all_files = [
        os.path.join(INPUT_DIR, f)
        for f in os.listdir(INPUT_DIR)
        if f.endswith(".xlsx") and f.startswith("Arbeitsmarkt-kommunal")
    ]

    result_frames = []
    for file in all_files:
        df = extract_fixed(file)
        result_frames.append(df)

    final_df = pd.concat(result_frames, ignore_index=True)

    output_path = os.path.join(OUTPUT_DIR, OUTPUT_FILENAME)
    final_df.to_excel(output_path, index=False)

    print(f"Result saved to {output_path}")

if __name__ == "__main__":
    main()
