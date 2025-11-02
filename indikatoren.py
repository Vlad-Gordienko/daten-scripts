# gemeinde; gemeinde_schluessel; jahr; variable; value;
# Bevölkerungsdichte (Einwohnerinnen auf 1 000 Einwohner); Bevölkerung (Insgesamt);


import os
import pandas as pd

from common.gebiet_schluessel import gebiet_schluessel
from common.mapping import get_gemeinde_from_gebiet, get_gemeinde_by_schluessel

AREAS = {
    6440001: 30.09,
    6440002: 32.54,
    6440003: 25.68,
    6440004: 122.88,
    6440005: 106.60,
    6440006: 37.65,
    6440007: 39.60,
    6440008: 50.18,
    6440009: 75.24,
    6440010: 12.67,
    6440011: 16.11,
    6440012: 43.94,
    6440013: 30.67,
    6440014: 12.50,
    6440015: 31.63,
    6440016: 118.33,
    6440017: 40.25,
    6440018: 37.65,
    6440019: 54.70,
    6440020: 34.26,
    6440021: 27.56,
    6440022: 16.14,
    6440023: 45.33,
    6440024: 43.11,
    6440025: 15.38,
}

FILENAME = "WK_Geburtsjahrgangsstatistik_Januar2025_KLW"
INPUT_DIR = "data"
OUTPUT_DIR = "result"
SHEET_NAME = "dadigesamt"
INPUT_FILENAME = os.path.join(INPUT_DIR, FILENAME + ".xlsx")
OUTPUT_FILENAME = os.path.join(OUTPUT_DIR, "indikatoren.csv")
YEARS = [2019, 2020, 2021, 2022, 2023, 2024]

def _make_base():
    xls = pd.ExcelFile(INPUT_FILENAME)
    df = pd.read_excel(xls, sheet_name=SHEET_NAME, dtype=str)
    df = df.dropna(subset=["Gebiet", "Jahrgang"])
    df["Gebiet"] = df["Gebiet"].str.strip()
    df["Jahrgang"] = pd.to_numeric(df["Jahrgang"], errors="coerce")
    df = df.dropna(subset=["Jahrgang"])
    df["Jahrgang"] = df["Jahrgang"].astype(int)
    if "EW gesamt" not in df.columns:
        return pd.DataFrame(columns=["Gemeinde","Jahrgang","EW gesamt"])
    df["EW gesamt"] = pd.to_numeric(df["EW gesamt"], errors="coerce").fillna(0).astype(int)
    df["Gemeinde"] = df["Gebiet"].map(get_gemeinde_from_gebiet)
    df = df[df["Gemeinde"].notnull() & (df["Gemeinde"] != "")]
    return df

def _build_for_year(base_df, year):
    def _grp(y):
        a = year - y
        if a < 21: return "junge"
        if a > 64: return "alte"
        return "mittleren"
    df = base_df[base_df["Jahrgang"] <= year].copy()
    df["gruppe"] = df["Jahrgang"].apply(_grp)
    g = df.groupby(["Gemeinde","gruppe"])["EW gesamt"].sum().unstack(fill_value=0).reset_index()
    g["gemeinde_schluessel"] = g["Gemeinde"].map(lambda x: gebiet_schluessel.get(x, ("",""))[0])
    g["gemeinde"] = g["gemeinde_schluessel"].apply(get_gemeinde_by_schluessel)
    g = g[~g["gemeinde"].isin(["Ausgewählte Gebiete zusammengefasst","Sanierungsgebiet"])]
    t = g["junge"] + g["mittleren"] + g["alte"]
    jq = ((g["junge"] / t).replace([float("inf"), -float("inf")], 0).fillna(0) * 100).round(2)
    aq = ((g["alte"]  / t).replace([float("inf"), -float("inf")], 0).fillna(0) * 100).round(2)
    fl = g["gemeinde_schluessel"].map(lambda s: AREAS.get(int(s), None) if str(s).isdigit() else None)
    out = pd.DataFrame({
        "gemeinde": g["gemeinde"],
        "gemeinde_schluessel": g["gemeinde_schluessel"],
        "jahr": year,
        "Jugendquotient": jq,
        "Altenquotiotent": aq,
        "Fläche": fl,
    })
    return out

def parse_excel():
    base = _make_base()
    parts = [_build_for_year(base, y) for y in YEARS]
    wide = pd.concat(parts, ignore_index=True)
    long = wide.melt(
        id_vars=["gemeinde","gemeinde_schluessel","jahr"],
        value_vars=["Altenquotiotent","Jugendquotient","Fläche"],
        var_name="variable",
        value_name="value"
    ).sort_values(["jahr","gemeinde","variable"], kind="mergesort")
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    long.to_csv(OUTPUT_FILENAME, index=False, sep=";")

if __name__ == "__main__":
    parse_excel()
