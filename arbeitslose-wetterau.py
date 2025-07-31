import pandas as pd
import os
import re

INPUT_DIR = "data/arbeitsortbeschäftigung"
OUTPUT_DIR = "result"
OUTPUT_FILENAME = "arbeitslose_wetterau.xlsx"

def extract_fixed(file_path: str) -> pd.DataFrame:
    df = pd.read_excel(file_path, sheet_name="Daten", header=None)

    years = [2020, 2021, 2022, 2023, 2024]

    row_total = 37
    row_male = 38
    row_female = 39
    row_sgb3 = 44
    row_sgb2 = 45

    values_total = pd.to_numeric(df.iloc[row_total, 2:7], errors="coerce").fillna(0).tolist()
    values_male = pd.to_numeric(df.iloc[row_male, 2:7], errors="coerce").fillna(0).tolist()
    values_female = pd.to_numeric(df.iloc[row_female, 2:7], errors="coerce").fillna(0).tolist()
    values_sgb3 = pd.to_numeric(df.iloc[row_sgb3, 2:7], errors="coerce").fillna(0).tolist()
    values_sgb2 = pd.to_numeric(df.iloc[row_sgb2, 2:7], errors="coerce").fillna(0).tolist()

    filename = os.path.basename(file_path)
    match = re.match(r"Arbeitsmarkt-kommunal_\d+_(.*)\.xlsx", filename)
    gemeinde = match.group(1).replace("_", " ") if match else "Unbekannt"

    rows = []
    for i, year in enumerate(years):
        row = {
            "Gemeinde": gemeinde,
            "Jahr": year,
            "Insgesamt": int(round(values_total[i])),
            "Männer": int(round(values_male[i])),
            "Frauen": int(round(values_female[i])),
            "SGB III": int(round(values_sgb3[i])),
            "SGB II": int(round(values_sgb2[i]))
        }
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

    sum_df = final_df.groupby("Jahr").sum(numeric_only=True).reset_index()
    sum_df["Gemeinde"] = "Wetteraukreis"

    columns = ["Gemeinde", "Jahr"] + [col for col in final_df.columns if col not in ["Gemeinde", "Jahr"]]
    sum_df = sum_df[columns]

    final_df = pd.concat([final_df, sum_df], ignore_index=True)

    output_path = os.path.join(OUTPUT_DIR, OUTPUT_FILENAME)
    final_df.to_excel(output_path, index=False)

    print(f"Result saved to {output_path}")

if __name__ == "__main__":
    main()
