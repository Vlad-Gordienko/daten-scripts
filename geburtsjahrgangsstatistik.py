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

def parse_excel_for_year(current_year):
    xls = pd.ExcelFile(INPUT_FILENAME)
    df = pd.read_excel(xls, sheet_name=SHEET_NAME, dtype=str)
    df = df.dropna(subset=["Gebiet", "Jahrgang"])
    df["Gebiet"] = df["Gebiet"].str.strip()
    df["Jahrgang"] = pd.to_numeric(df["Jahrgang"], errors='coerce')
    df = df.dropna(subset=["Jahrgang"])
    df["Jahrgang"] = df["Jahrgang"].astype(int)
    df = df[df["Jahrgang"] <= current_year]

    df["Gemeinde"] = df["Gebiet"].map(get_gemeinde_from_gebiet)
    all_gebieten = df["Gebiet"].dropna().unique().tolist()
    undetected_gebiete = track_undetected_gebiete(all_gebieten)
    log_missing_gebiete(undetected_gebiete)

    if "EW gesamt" not in df.columns:
        return pd.DataFrame()

    df["EW gesamt"] = pd.to_numeric(df["EW gesamt"], errors='coerce').fillna(0).astype(int)

    def classify_group(year):
        age = current_year - year
        if age < 21:
            return "junge"
        elif age > 64:
            return "alte"
        return "mittleren"

    df["gruppe"] = df["Jahrgang"].apply(classify_group)
    df = df[df["Gemeinde"].notnull() & (df["Gemeinde"] != "")]

    grouped = df.groupby(["Gemeinde", "gruppe"])["EW gesamt"].sum().unstack(fill_value=0).reset_index()
    grouped["gemeinde_schluessel"] = grouped["Gemeinde"].map(lambda x: gebiet_schluessel.get(x, ("", ""))[0])
    grouped["gemeinde_schluessel"] = grouped["gemeinde_schluessel"].apply(
        lambda x: str(x).zfill(8) if pd.notnull(x) and str(x).isdigit() else ""
    )
    grouped["gemeinde"] = grouped["gemeinde_schluessel"].apply(get_gemeinde_by_schluessel)
    grouped["iso"] = grouped["Gemeinde"].map(lambda x: gebiet_schluessel.get(x, ("", ""))[1])

    grouped["junge_quotient"] = (grouped["junge"] / grouped["mittleren"]).replace([float("inf"), -float("inf")], 0) * 100
    grouped["alte_quotient"] = (grouped["alte"] / grouped["mittleren"]).replace([float("inf"), -float("inf")], 0) * 100
    grouped["junge_quotient"] = grouped["junge_quotient"].round(2).astype(str)
    grouped["alte_quotient"] = grouped["alte_quotient"].round(2).astype(str)

    grouped = grouped[~grouped["gemeinde"].isin(["Ausgewählte Gebiete zusammengefasst", "Sanierungsgebiet"])]
    grouped["jahr"] = 2024

    final_columns = [
        "gemeinde",
        "gemeinde_schluessel",
        "iso",
        "junge",
        "alte",
        "mittleren",
        "junge_quotient",
        "alte_quotient",
        "jahr"
    ]
    return grouped[final_columns]


def add_summary_row(df):
    summary_rows = []
    for year in df["jahr"].unique():
        subset = df[df["jahr"] == year]
        junge_sum = subset["junge"].sum()
        alte_sum = subset["alte"].sum()
        mittleren_sum = subset["mittleren"].sum()
        junge_quot = f"{round((junge_sum / mittleren_sum) * 100, 2)}" if mittleren_sum else "0"
        alte_quot = f"{round((alte_sum / mittleren_sum) * 100, 2)}" if mittleren_sum else "0"

        summary_rows.append({
            "gemeinde": "Wetteraukreis",
            "gemeinde_schluessel": "06440000",
            "iso": "0",
            "junge": junge_sum,
            "alte": alte_sum,
            "mittleren": mittleren_sum,
            "junge_quotient": junge_quot,
            "alte_quotient": alte_quot,
            "jahr": year
        })

    return pd.concat([df, pd.DataFrame(summary_rows)], ignore_index=True)


def reorder_with_sum_after_each_year(df):
    ordered_rows = []
    for year in sorted(df["jahr"].unique()):
        part = df[df["jahr"] == year]
        summary = part[part["gemeinde"] == "SUMME"]
        rest = part[part["gemeinde"] != "SUMME"]
        ordered_rows.append(pd.concat([rest, summary], ignore_index=True))
    return pd.concat(ordered_rows, ignore_index=True)


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
    YEARS = [2019, 2020, 2021, 2022, 2023, 2024]

    results = [parse_excel_for_year(year) for year in YEARS]
    combined = pd.concat(results, ignore_index=True)
    with_sum = add_summary_row(combined)
    final = reorder_with_sum_after_each_year(with_sum)

    final.to_excel(OUTPUT_FILENAME, index=False)
    print(f"Result saved to {OUTPUT_FILENAME}")


if __name__ == "__main__":
    parse_excel()
