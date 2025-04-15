import pandas as pd
import os

FILENAME = "arbeitsortbeschäftigung.xlsx"
INPUT_DIR = "data/arbeitsortbeschäftigung"
OUTPUT_DIR = "result"
SHEET_NAME = "Daten"

INPUT_PATH = os.path.join(INPUT_DIR)
OUTPUT_PATH = os.path.join(OUTPUT_DIR, FILENAME)


def parse_value(v):
    try:
        val = float(v)
        return val
    except:
        return -1

def parse_arbeitsmarkt():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    all_files = [f for f in os.listdir(INPUT_PATH) if f.endswith(".xlsx") and f.startswith("Arbeitsmarkt-kommunal")]
    if not all_files:
        print(f"No Arbeitsmarkt files found in '{INPUT_PATH}'")
        return

    data_rows = []

    for file in sorted(all_files):
        full_path = os.path.join(INPUT_PATH, file)

        try:
            df = pd.read_excel(full_path, sheet_name=SHEET_NAME, header=None, usecols="B,G", skiprows=17, nrows=4)
            category_names = df.iloc[:, 0].tolist()
            values = [parse_value(v) for v in df.iloc[:, 1].tolist()]

            filename_core = os.path.splitext(file)[0]
            parts = filename_core.split("_")
            commune_raw = "_".join(parts[2:])

            commune_name = (
                commune_raw.replace("_Stadt", "")
                           .replace("_Gemeinde", "")
                           .replace("_Stadtteil", "")
                           .replace("_Landkreis", "")
                           .replace("_Kreis", "")
                           .replace("_", " ")
                           .strip()
            )

            row = {"gemeinde": commune_name}
            for i in range(4):
                row[category_names[i]] = values[i]

            data_rows.append(row)

        except Exception as e:
            print(f"Error processing {file}: {e}")

    if not data_rows:
        print("No valid data extracted.")
        return

    result_df = pd.DataFrame(data_rows)
    result_df.to_excel(OUTPUT_PATH, index=False)
    print(f"Result saved to {OUTPUT_PATH}")


if __name__ == "__main__":
    parse_arbeitsmarkt()
