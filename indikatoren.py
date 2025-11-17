import os
import pandas as pd

from common.gebiet_schluessel import gebiet_schluessel
from common.mapping import get_gemeinde_from_gebiet
from common.schluessel_map import schluessel_map

# Fest hinterlegte Flächen (km^2)
AREAS = {
    6440001: 30.09, 6440002: 32.54, 6440003: 25.68, 6440004: 122.88, 6440005: 106.60,
    6440006: 37.65, 6440007: 39.60, 6440008: 50.18, 6440009: 75.24, 6440010: 12.67,
    6440011: 16.11, 6440012: 43.94, 6440013: 30.67, 6440014: 12.50, 6440015: 31.63,
    6440016: 118.33, 6440017: 40.25, 6440018: 37.65, 6440019: 54.70, 6440020: 34.26,
    6440021: 27.56, 6440022: 16.14, 6440023: 45.33, 6440024: 43.11, 6440025: 15.38,
}

# Ein-/Ausgabe
FILENAME = "WK_Geburtsjahrgangsstatistik_Januar2025_KLW"
INPUT_DIR = "data"
OUTPUT_DIR = "result"
SHEET_NAME = "dadigesamt"
INPUT_FILENAME = os.path.join(INPUT_DIR, FILENAME + ".xlsx")
OUTPUT_FILENAME = os.path.join(OUTPUT_DIR, "indikatoren.csv")
YEARS = [2019, 2020, 2021, 2022, 2023, 2024]


def _make_base() -> pd.DataFrame:
    """
    Haupttabelle laden und Basisdaten vorbereiten:
    - Gemeinde
    - Jahrgang
    - EW gesamt
    """
    xls = pd.ExcelFile(INPUT_FILENAME)
    df = pd.read_excel(xls, sheet_name=SHEET_NAME, dtype=str)

    # Pflichtspalten
    df = df.dropna(subset=["Gebiet", "Jahrgang"])
    df["Gebiet"] = df["Gebiet"].str.strip()

    # Jahrgang in int
    df["Jahrgang"] = pd.to_numeric(df["Jahrgang"], errors="coerce")
    df = df.dropna(subset=["Jahrgang"])
    df["Jahrgang"] = df["Jahrgang"].astype(int)

    # Ohne EW gesamt – ничего не посчитаем
    if "EW gesamt" not in df.columns:
        return pd.DataFrame(columns=["Gemeinde", "Jahrgang", "EW gesamt"])

    # Einwohner gesamt in int, мусор → 0
    df["EW gesamt"] = pd.to_numeric(df["EW gesamt"], errors="coerce").fillna(0).astype(int)

    # Gemeinde zuordnen
    df["Gemeinde"] = df["Gebiet"].map(get_gemeinde_from_gebiet)
    df = df[df["Gemeinde"].notnull() & (df["Gemeinde"] != "")]

    return df


def _build_for_year(base_df: pd.DataFrame, year: int) -> pd.DataFrame:
    """
    Für ein Zieljahr:
    - Einwohner nach Altersgruppen aufsummieren
    - Jugend-/Altenquotient berechnen
    - Bevölkerung gesamt berechnen
    - Fläche zuordnen
    - Bevölkerungsdichte berechnen
    """
    def _grp(yahrgang: int) -> str:
        alter = year - yahrgang
        if alter < 21:
            return "junge"
        if alter > 64:
            return "alte"
        return "mittleren"

    # Nur Jahrgänge <= Jahr betrachten
    df = base_df[base_df["Jahrgang"] <= year].copy()
    df["gruppe"] = df["Jahrgang"].apply(_grp)

    # Summe EW gesamt pro Gemeinde und Altersgruppe
    g = (
        df.groupby(["Gemeinde", "gruppe"])["EW gesamt"]
        .sum()
        .unstack(fill_value=0)
        .reset_index()
    )

    # Schlüssel + saubere Gemeindebezeichnung
    g["gemeinde_schluessel"] = g["Gemeinde"].map(
        lambda x: str(gebiet_schluessel.get(x, ("", ""))[0])
    )
    g["gemeinde"] = g["gemeinde_schluessel"].map(
        lambda k: schluessel_map.get(int(k)) if k.isdigit() else None
    )
    g = g[g["gemeinde"].notna()]

    # Gesamtbevölkerung im Jahr
    t = g["junge"] + g["mittleren"] + g["alte"]
    bevoelkerung = t

    # Quotienten (%)
    jq = (
        (g["junge"] / t)
        .replace([float("inf"), -float("inf")], 0)
        .fillna(0)
        * 100
    ).round(2)

    aq = (
        (g["alte"] / t)
        .replace([float("inf"), -float("inf")], 0)
        .fillna(0)
        * 100
    ).round(2)

    # Fläche aus AREAS
    fl = g["gemeinde_schluessel"].map(
        lambda s: AREAS.get(int(s), None) if str(s).isdigit() else None
    )
    fl_num = pd.to_numeric(fl, errors="coerce")

    # Bevölkerungsdichte (Einw./km²)
    bevoelkerungsdichte = (bevoelkerung / fl_num)
    bevoelkerungsdichte = bevoelkerungsdichte.replace(
        [float("inf"), -float("inf")], pd.NA
    ).round(0)

    return pd.DataFrame({
        "gemeinde": g["gemeinde"],
        "gemeinde_schluessel": g["gemeinde_schluessel"],
        "jahr": year,
        "Jugendquotient": jq,
        "Altenquotiotent": aq,
        "Fläche": fl_num,
        "Bevölkerung": bevoelkerung,
        "Bevölkerungsdichte": bevoelkerungsdichte,
    })

def parse_excel():
    base = _make_base()

    if base.empty:
        combined = pd.DataFrame(
            columns=[
                "gemeinde",
                "gemeinde_schluessel",
                "jahr",
                "Jugendquotient",
                "Altenquotiotent",
                "Fläche",
                "Bevölkerung",
                "Bevölkerungsdichte",
            ]
        )
    else:
        parts = [_build_for_year(base, y) for y in YEARS]
        combined = pd.concat(parts, ignore_index=True)

    # Long-Format für alle Indikatoren
    indikatoren = combined.melt(
        id_vars=["gemeinde", "gemeinde_schluessel", "jahr"],
        value_vars=[
            "Altenquotiotent",
            "Jugendquotient",
            "Fläche",
            "Bevölkerung",
            "Bevölkerungsdichte",
        ],
        var_name="variable",
        value_name="value",
    )

    # Falls Name fehlt, aus Schlüssel ergänzen
    if indikatoren["gemeinde"].isna().any():
        code2name = {str(k): v for k, v in schluessel_map.items()}
        indikatoren["gemeinde"] = indikatoren["gemeinde"].fillna(
            indikatoren["gemeinde_schluessel"].map(code2name)
        )

    # Sortieren und schreiben
    indikatoren = indikatoren.sort_values(
        ["jahr", "gemeinde_schluessel", "variable"]
    ).reset_index(drop=True)

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    indikatoren.to_csv(OUTPUT_FILENAME, index=False, sep=";")


if __name__ == "__main__":
    parse_excel()
