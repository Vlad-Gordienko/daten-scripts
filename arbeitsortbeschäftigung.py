import pandas as pd
import os

from common.mapping import normalize_gemeinde_name

FILENAME = "arbeitsortbeschäftigung.xlsx"
INPUT_DIR = "data/arbeitsortbeschäftigung"
OUTPUT_DIR = "result"
SHEET_NAME = "Daten"

INPUT_PATH = os.path.join(INPUT_DIR)
OUTPUT_PATH = os.path.join(OUTPUT_DIR, FILENAME)

def parse_value(v):
    """
    Converts input to float if possible, otherwise returns 0.
    Used to sanitize Excel values that may be empty or contain symbols like '*'.
    """
    try:
        val = float(v)
        return val
    except:
        return 0

def parse_arbeitsmarkt():
    """
    Parses a batch of Excel files with Arbeitsmarkt-Kommunal data.
    - Extracts specific metrics from sheet 'Daten'
    - Normalizes Gemeinde names
    - Aggregates data into a single Excel file
    """
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # Filter all Arbeitsmarkt-related Excel files
    all_files = [f for f in os.listdir(INPUT_PATH) if f.endswith(".xlsx") and f.startswith("Arbeitsmarkt-kommunal")]
    if not all_files:
        print(f"No Arbeitsmarkt files found in '{INPUT_PATH}'")
        return

    data_rows = []

    for file in sorted(all_files):
        full_path = os.path.join(INPUT_PATH, file)

        try:
            # Read 4 rows from columns B and G; they contain metric names and values
            df = pd.read_excel(full_path, sheet_name=SHEET_NAME, header=None, usecols="B,G", skiprows=17, nrows=4)
            category_names = df.iloc[:, 0].tolist()
            values = [parse_value(v) for v in df.iloc[:, 1].tolist()]

            # Extract and clean the raw Gemeinde name from filename
            filename_core = os.path.splitext(file)[0]
            parts = filename_core.split("_")
            commune_raw = "_".join(parts[2:])

            # Remove suffixes like "_Stadt" and normalize formatting
            commune_name = (
                commune_raw.replace("_Stadt", "")
                           .replace("_Gemeinde", "")
                           .replace("_Stadtteil", "")
                           .replace("_Landkreis", "")
                           .replace("_Kreis", "")
                           .replace("_", " ")
                           .strip()
            )

            # Use standardized/canonical name
            row = {"gemeinde": normalize_gemeinde_name(commune_name)}
            for i in range(4):
                row[category_names[i]] = values[i]

            data_rows.append(row)

        except Exception as e:
            print(f"Error processing {file}: {e}")

    if not data_rows:
        print("No valid data extracted.")
        return

    # Build final table and export to Excel
    result_df = pd.DataFrame(data_rows)
    result_df.to_excel(OUTPUT_PATH, index=False)
    print(f"Result saved to {OUTPUT_PATH}")

if __name__ == "__main__":
    parse_arbeitsmarkt()
