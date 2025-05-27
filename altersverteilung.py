import pandas as pd
import os
import re

INPUT_DIR = "data/altersverteilung"
OUTPUT_DIR = "result"
OUTPUT_FILENAME = "altersstruktur_wetterau.xlsx"

def extract_summary_by_year(file_path: str) -> dict:
    """
    Reads a yearly Excel file and extracts total population counts
    grouped by age category: 0–20, 21–64, and 65+.

    Assumes:
    - File name starts with the year (e.g. "2023 GjS Wetteraukreis mit GKZ.XLSX")
    - Sheet name matches pattern "<year> GjS Wetteraukreis"
    - Columns 'Jahrgang' and 'EW gesamt' are present
    """

    filename = os.path.basename(file_path)
    year = int(filename.split()[0])  # Extract year from filename
    sheet_name = f"{year} GjS Wetteraukreis"

    # Load only required columns
    df = pd.read_excel(file_path, sheet_name=sheet_name, usecols=["Jahrgang", "EW gesamt"])
    df = df.dropna(subset=["Jahrgang", "EW gesamt"])
    df["Jahrgang"] = pd.to_numeric(df["Jahrgang"], errors="coerce")
    df["EW gesamt"] = pd.to_numeric(df["EW gesamt"], errors="coerce")
    df = df.dropna()

    # Calculate actual age based on current year
    df["Jahrgang"] = df["Jahrgang"].astype(int)
    df["EW gesamt"] = df["EW gesamt"].astype(int)
    df["Alter"] = year - df["Jahrgang"]

    # Group population into age buckets
    count_0_20 = df[df["Alter"] <= 20]["EW gesamt"].sum()
    count_21_64 = df[(df["Alter"] >= 21) & (df["Alter"] <= 64)]["EW gesamt"].sum()
    count_65_plus = df[df["Alter"] >= 65]["EW gesamt"].sum()

    return {
        "Jahr": year,
        "0 - 20": count_0_20,
        "21 - 64": count_21_64,
        "65+": count_65_plus
    }

def main():
    """
    Main execution:
    - Searches for all matching Excel files in the input folder
    - Extracts demographic summaries from each
    - Combines and exports as a single Excel summary
    """
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # Match files like "2023 GjS Wetteraukreis mit GKZ.XLSX"
    all_files = [
        os.path.join(INPUT_DIR, f)
        for f in os.listdir(INPUT_DIR)
        if re.match(r"\d{4} GjS Wetteraukreis mit GKZ\.XLSX$", f)
    ]

    if not all_files:
        print(f"No valid files found in '{INPUT_DIR}'")
        return

    # Process each file and combine results
    summary_data = [extract_summary_by_year(path) for path in all_files]
    summary_df = pd.DataFrame(summary_data).sort_values(by="Jahr")

    output_path = os.path.join(OUTPUT_DIR, OUTPUT_FILENAME)
    summary_df.to_excel(output_path, index=False)

    print(f"Result saved to {output_path}")

if __name__ == "__main__":
    main()
