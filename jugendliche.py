import os
import pandas as pd

from common.mapping import get_gemeinde_from_gebiet
from common.schluessel_map import schluessel_map

FILENAME = "WK_Geburtsjahrgangsstatistik_Januar2025_KLW"
INPUT_DIR = "data"
OUTPUT_DIR = "result"
SHEET_NAME = "dadigesamt"
INPUT_FILENAME = os.path.join(INPUT_DIR, FILENAME + ".xlsx")
OUTPUT_FILENAME = os.path.join(OUTPUT_DIR, "jugendliche.csv")
YEARS = [2019, 2020, 2021, 2022, 2023, 2024]

NAME2CODE = {v: str(k) for k, v in schluessel_map.items()}
CODE2NAME = {str(k): v for k, v in schluessel_map.items()}

def load_base():
    df = pd.read_excel(INPUT_FILENAME, sheet_name=SHEET_NAME, dtype=str)
    df = df.dropna(subset=["Gebiet", "Jahrgang"])
    df["Gemeinde"] = df["Gebiet"].str.strip().map(get_gemeinde_from_gebiet)
    df = df[df["Gemeinde"].notna() & (df["Gemeinde"] != "")]
    df["Jahrgang"] = pd.to_numeric(df["Jahrgang"], errors="coerce")
    df = df.dropna(subset=["Jahrgang"])
    df["Jahrgang"] = df["Jahrgang"].astype(int)
    df["EW gesamt"] = pd.to_numeric(df["EW gesamt"], errors="coerce").fillna(0).astype(int)
    return df[["Gemeinde", "Jahrgang", "EW gesamt"]]

def build_year(df, year):
    d = df.copy()
    d["age"] = year - d["Jahrgang"]
    d = d[(d["age"] >= 0) & (d["age"] <= 21)]
    d["variable"] = pd.cut(d["age"], bins=[-1, 6, 14, 21], labels=["0-6", "7-14", "15-21"])
    agg = (
        d.groupby(["Gemeinde", "variable"], observed=False)["EW gesamt"]
         .sum()
         .unstack(fill_value=0)
         .reset_index()
    )
    agg["Gemeinde"] = agg["Gemeinde"].astype(str).str.strip()
    is_code = agg["Gemeinde"].str.fullmatch(r"\d{7}")
    agg.loc[is_code, "gemeinde_schluessel"] = agg.loc[is_code, "Gemeinde"]
    agg.loc[~is_code, "gemeinde_schluessel"] = agg.loc[~is_code, "Gemeinde"].map(NAME2CODE)
    agg.loc[is_code, "gemeinde"] = agg.loc[is_code, "Gemeinde"].map(CODE2NAME)
    agg.loc[~is_code, "gemeinde"] = agg.loc[~is_code, "Gemeinde"]
    agg = agg[agg["gemeinde"].notna()]
    agg["jahr"] = year
    long = agg.melt(
        id_vars=["gemeinde","gemeinde_schluessel","jahr"],
        value_vars=["0-6","7-14","15-21"],
        var_name="variable",
        value_name="value",
    )
    return long[["gemeinde","gemeinde_schluessel","jahr","variable","value"]]

def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    base = load_base()
    parts = [build_year(base, y) for y in YEARS]
    result = pd.concat(parts, ignore_index=True)
    result.to_csv(OUTPUT_FILENAME, index=False, sep=";")

if __name__ == "__main__":
    main()
