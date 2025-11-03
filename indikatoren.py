import os
import re
import glob
import warnings
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


# ------------------------------------------------------------
# Teil 1: Basisindikatoren aus Jahrgangstabelle (Jugend/Alten/Fläche)
# ------------------------------------------------------------

def _make_base() -> pd.DataFrame:
    """Haupttabelle laden und Basisdaten vorbereiten (Gemeinde, Jahrgang, EW gesamt)."""
    xls = pd.ExcelFile(INPUT_FILENAME)
    df = pd.read_excel(xls, sheet_name=SHEET_NAME, dtype=str)

    df = df.dropna(subset=["Gebiet", "Jahrgang"])
    df["Gebiet"] = df["Gebiet"].str.strip()
    df["Jahrgang"] = pd.to_numeric(df["Jahrgang"], errors="coerce")
    df = df.dropna(subset=["Jahrgang"])
    df["Jahrgang"] = df["Jahrgang"].astype(int)

    if "EW gesamt" not in df.columns:
        return pd.DataFrame(columns=["Gemeinde", "Jahrgang", "EW gesamt"])

    df["EW gesamt"] = pd.to_numeric(df["EW gesamt"], errors="coerce").fillna(0).astype(int)
    df["Gemeinde"] = df["Gebiet"].map(get_gemeinde_from_gebiet)
    df = df[df["Gemeinde"].notnull() & (df["Gemeinde"] != "")]
    return df


def _build_for_year(base_df: pd.DataFrame, year: int) -> pd.DataFrame:
    """Für ein Zieljahr Gruppen bilden, summieren und Jugend-/Altenquotient sowie Fläche ableiten."""
    def _grp(y):
        a = year - y
        if a < 21:
            return "junge"
        if a > 64:
            return "alte"
        return "mittleren"

    df = base_df[base_df["Jahrgang"] <= year].copy()
    df["gruppe"] = df["Jahrgang"].apply(_grp)

    g = (
        df.groupby(["Gemeinde", "gruppe"])["EW gesamt"]
          .sum().unstack(fill_value=0).reset_index()
    )

    # Schlüssel + saubere Gemeindebezeichnung
    g["gemeinde_schluessel"] = g["Gemeinde"].map(lambda x: str(gebiet_schluessel.get(x, ("", ""))[0]))
    g["gemeinde"] = g["gemeinde_schluessel"].map(lambda k: schluessel_map.get(int(k)) if k.isdigit() else None)
    g = g[g["gemeinde"].notna()]

    # Quoten (%)
    t = g["junge"] + g["mittleren"] + g["alte"]
    jq = ((g["junge"] / t).replace([float("inf"), -float("inf")], 0).fillna(0) * 100).round(2)
    aq = ((g["alte"]  / t).replace([float("inf"), -float("inf")], 0).fillna(0) * 100).round(2)

    # Fläche aus AREAS
    fl = g["gemeinde_schluessel"].map(lambda s: AREAS.get(int(s), None) if str(s).isdigit() else None)

    return pd.DataFrame({
        "gemeinde": g["gemeinde"],
        "gemeinde_schluessel": g["gemeinde_schluessel"],
        "jahr": year,
        "Jugendquotient": jq,
        "Altenquotiotent": aq,
        "Fläche": fl,
    })


# ------------------------------------------------------------
# Teil 2: HGST-Dateien (Bevölkerung/Insgesamt, Bevölkerungsdichte)
# ------------------------------------------------------------

def _to_number(s):
    """Zahlrobustheit: Tausender, NBSP, Komma/Dezimalpunkt."""
    if s is None:
        return None
    v = str(s).strip().replace("\u00a0", " ")        # NBSP -> Space
    v = re.sub(r"\s+", "", v)                        # alle Leerzeichen raus
    v = v.replace(",", ".")                          # Komma -> Punkt
    v = re.sub(r"\.(?=\d{3}(\D|$))", "", v)          # Punkte vor 3er-Gruppen entfernen
    try:
        return float(v)
    except:
        return None


def _normalize_ags(s):
    """AGS aus HGST normalisieren auf 7-stellig (z. B. 440xyz -> 6440xyz)."""
    if pd.isna(s):
        return None
    s = str(s).replace(" ", "").replace(".", "")
    if not re.fullmatch(r"\d+", s):
        return None
    s2 = s.lstrip("0")
    if len(s2) == 6 and s2.startswith("440"):   # 440xxx -> 6440xxx
        return "6" + s2
    if len(s2) == 8 and s2.startswith("644000"):  # 8-stellig -> 7-stellig
        return s2[1:]
    if len(s2) == 7:
        return s2
    return s2


def _pick_sheet(path):
    """Passendes Tabellenblatt je nach Jahr/Datei auswählen (Namensvarianten)."""
    xls = pd.ExcelFile(path)
    sheets = xls.sheet_names
    for cand in ["Bevölkerung und Gebiet", "Bevölkerung", "1"]:
        if cand in sheets:
            return cand
    return sheets[0]


def _read_hgst_file(path, year) -> pd.DataFrame:
    """Eine HGST-Datei lesen und lange Form (Bevölkerung, Bevölkerungsdichte) für Wetterau liefern."""
    # openpyxl-UserWarning zum Druckbereich nur lokal unterdrücken
    with warnings.catch_warnings():
        warnings.filterwarnings(
            "ignore",
            message=r"Print area cannot be set to Defined name",
            category=UserWarning,
            module=r"openpyxl\.reader\.workbook"
        )
        sheet = _pick_sheet(path)
        raw = pd.read_excel(path, sheet_name=sheet, header=None, dtype=str)

    # Soft-Cleaning
    raw = raw.apply(lambda col: col.map(
        lambda x: x.replace("\u00a0", " ").replace("\n", " ").strip() if isinstance(x, str) else x
    ))

    # Headerzeile heuristisch finden
    header_idx = 0
    for i in range(min(25, len(raw))):
        row = raw.iloc[i].astype(str).str.lower()
        if any(c.strip() == "gebiet" for c in row) and any("schlüssel" in c for c in row):
            header_idx = i
            break

    # Header zusammenbauen (ggf. Zeile darunter mergen)
    top = raw.iloc[header_idx].tolist()
    nxt = raw.iloc[header_idx + 1].tolist() if header_idx + 1 < len(raw) else [None] * len(top)

    def comb(a, b):
        a = "" if a is None or str(a) == "nan" else str(a).strip()
        b = "" if b is None or str(b) == "nan" else str(b).strip()
        if a and b:
            return f"{a} {b}"
        return a or b or "nan"

    headers = [comb(a, b) for a, b in zip(top, nxt)]
    has_metrics = any(("insgesamt" in str(h).lower()) or ("je km" in str(h).lower()) or ("1 000" in str(h)) for h in headers)
    if not has_metrics:
        headers = [str(x).strip() if x is not None else "nan" for x in top]

    body = raw.iloc[header_idx + 1:].reset_index(drop=True)
    body.columns = [str(c).strip() for c in headers]

    # Relevante Spalten identifizieren
    cols = body.columns.tolist()
    low = [c.lower() for c in cols]
    ags  = next((cols[i] for i, c in enumerate(low) if "schlüssel" in c and "amtlicher" in c), None) \
        or next((cols[i] for i, c in enumerate(low) if "schlüssel" in c), None)
    name = next((cols[i] for i, c in enumerate(low) if c.strip() == "gebiet"), None)
    bev  = next((cols[i] for i, c in enumerate(low) if c.endswith(" insgesamt") or " insgesamt" in c or c.strip() == "insgesamt"), None)
    dens = next((cols[i] for i, c in enumerate(low) if "einwohnerinnen auf 1 000 einwohner" in c), None)

    if not all([ags, name, bev, dens]):
        return pd.DataFrame(columns=["gemeinde", "gemeinde_schluessel", "jahr", "variable", "value"])

    work = body[[ags, name, bev, dens]].copy()
    work.columns = ["ags", "gebiet", "bev", "dens"]
    work = work.dropna(subset=["ags"])
    work["ags_norm"] = work["ags"].map(_normalize_ags)

    # Wetteraukreis: nur 6440xxx (ohne Kreissumme 6440000)
    wett = work[work["ags_norm"].astype(str).str.startswith("6440")].copy()
    wett = wett[wett["ags_norm"] != "6440000"]

    # Zahlen parsen/formatieren
    wett["Bevölkerung"] = wett["bev"].map(_to_number).round(0).astype("Int64")
    wett["Bevölkerungsdichte"] = wett["dens"].map(_to_number).round(0).astype("Int64")

    # Gemeindename via Schlüssel
    wett["gemeinde_schluessel"] = wett["ags_norm"].astype(str)
    code2name = {str(k): v for k, v in schluessel_map.items()}
    wett["gemeinde"] = wett["gemeinde_schluessel"].map(code2name)
    wett["jahr"] = int(year)

    wide = wett[["gemeinde", "gemeinde_schluessel", "jahr", "Bevölkerung", "Bevölkerungsdichte"]].copy()
    return wide.melt(
        id_vars=["gemeinde", "gemeinde_schluessel", "jahr"],
        value_vars=["Bevölkerung", "Bevölkerungsdichte"],
        var_name="variable",
        value_name="value"
    )


def _build_hgst_all() -> pd.DataFrame:
    """Alle HGST-Jahresdateien im Eingabeordner lesen und zusammenführen."""
    files = sorted(glob.glob(os.path.join(INPUT_DIR, "HGSt_j*.xlsx")))
    parts = []
    for f in files:
        m = re.search(r"(\d{4})", os.path.basename(f))
        y = int(m.group(1)) if m else None
        parts.append(_read_hgst_file(f, y))
    if not parts:
        return pd.DataFrame(columns=["gemeinde", "gemeinde_schluessel", "jahr", "variable", "value"])
    return pd.concat(parts, ignore_index=True)


# ------------------------------------------------------------
# Hauptablauf
# ------------------------------------------------------------

def parse_excel():
    # Basisindikatoren (Jugend/Alten/Fläche) pro Jahr
    base = _make_base()
    parts = [_build_for_year(base, y) for y in YEARS]
    combined = pd.concat(parts, ignore_index=True)

    left = combined.melt(
        id_vars=["gemeinde", "gemeinde_schluessel", "jahr"],
        value_vars=["Altenquotiotent", "Jugendquotient", "Fläche"],
        var_name="variable",
        value_name="value"
    )

    # HGST-Indikatoren (Bevölkerung, Bevölkerungsdichte)
    right = _build_hgst_all()

    # Falls Name fehlt, aus Schlüssel ergänzen
    if left["gemeinde"].isna().any():
        code2name = {str(k): v for k, v in schluessel_map.items()}
        left["gemeinde"] = left["gemeinde"].fillna(left["gemeinde_schluessel"].map(code2name))

    # Zusammenführen, sortieren, schreiben
    out = pd.concat([left, right], ignore_index=True)
    out = out.sort_values(["jahr", "gemeinde_schluessel", "variable"]).reset_index(drop=True)

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    out.to_csv(OUTPUT_FILENAME, index=False, sep=";")


if __name__ == "__main__":
    parse_excel()
