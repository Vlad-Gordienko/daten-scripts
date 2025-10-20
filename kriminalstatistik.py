# -*- coding: utf-8 -*-
"""
- Alle xlsx in data/kriminalstatistik/ lesen
- Nur Wetteraukreis filtern
- Spalten mappen und Zahlen normalisieren
- 'jahr' setzen
- Ein CSV schreiben (UTF-8, ';')
- Datenkatalog (Header-Mapping) als CSV exportieren
"""

from pathlib import Path
import pandas as pd
import math
import re

# --- Pfade / Einstellungen
BASE_DIR = Path("data/kriminalstatistik")   # Ordner mit xlsx
SHEET_NAME = "T01_Kreise"                   # Blattname
OUT_CSV = BASE_DIR / "kriminalstatistik_wetteraukreis_2020_2024.csv"
OUT_DICT = BASE_DIR / "datenkatalog_kriminalstatistik_wetteraukreis.csv"

# --- Hilfen
def find_header_row(xlsx_path: Path, sheet: str) -> int:
    """Finde Zeile mit 'Schlüssel' und 'Straftat'."""
    df_raw = pd.read_excel(xlsx_path, sheet_name=sheet, header=None)
    for i in range(min(60, len(df_raw))):
        row_vals = df_raw.iloc[i].astype(str).str.strip().tolist()
        if "Schlüssel" in row_vals and "Straftat" in row_vals:
            return i
    raise ValueError(f"Header-Zeile nicht gefunden: {xlsx_path.name}")

def is_unnamed(x) -> bool:
    """Ist leer oder 'Unnamed'?"""
    if x is None or (isinstance(x, float) and math.isnan(x)):
        return True
    s = str(x).strip()
    return (s == "") or s.lower().startswith("unnamed")

# Untertitel-Reihenfolgen (Original-Namen unten)
GROUP_SPECS = {
    "erfasste Fälle davon: Versuche": ["Anzahl", "in %"],
    "mit Schusswaffe": ["gedroht", "geschossen"],
    "Aufklärung": ["Anzahl Fälle", "AQ"],
    "Tatverdächtige": ["insgesamt", "männlich", "weiblich"],
    "Nichtdeutsche Tatverdächtige": ["Anzahl", "Anteil an TV insg. in %"],
}

# Mapping RAW -> final (Original rechts im Kommentar lassen)
RAW_TO_FINAL = {
    ("Schlüssel", ""): "schluessel",                            # "Schlüssel"
    ("Straftat", ""): "straftat",                               # "Straftat"
    ("Kreisart", ""): "kreisart",                               # "Kreisart"
    ("Anzahl erfasste Fälle", ""): "faelle",                    # "Anzahl erfasste Fälle"
    ("HZ", ""): "hz_pro_100k",                                  # "HZ"
    ("HZ nach Zensus 2022", ""): "hz_pro_100k",                 # "HZ nach Zensus 2022"
    ("erfasste Fälle davon: Versuche", "Anzahl"): "versuche_anzahl",           # "… Versuche" -> "Anzahl"
    ("erfasste Fälle davon: Versuche", "in %"): "versuche_prozent",            # "… Versuche" -> "in %"
    ("mit Schusswaffe", "gedroht"): "mit_schusswaffe_gedroht",                # "mit Schusswaffe" -> "gedroht"
    ("mit Schusswaffe", "geschossen"): "mit_schusswaffe_geschossen",          # "mit Schusswaffe" -> "geschossen"
    ("Aufklärung", "Anzahl Fälle"): "aufklaerung_faelle",                     # "Aufklärung" -> "Anzahl Fälle"
    ("Aufklärung", "AQ"): "aq_prozent",                                       # "Aufklärung" -> "AQ"
    ("Tatverdächtige", "insgesamt"): "tatverdaechtige_insgesamt",             # "Tatverdächtige" -> "insgesamt"
    ("Tatverdächtige", "männlich"): "tatverdaechtige_maennlich",              # "Tatverdächtige" -> "männlich"
    ("Tatverdächtige", "weiblich"): "tatverdaechtige_weiblich",               # "Tatverdächtige" -> "weiblich"
    ("Nichtdeutsche Tatverdächtige", "Anzahl"): "nichtdeutsche_tv_anzahl",    # "Nichtdeutsche …" -> "Anzahl"
    ("Nichtdeutsche Tatverdächtige", "Anteil an TV insg. in %"): "nichtdeutsche_tv_anteil_prozent",  # "Nichtdeutsche …" -> "Anteil an TV insg. in %"
}

FINAL_ORDER = [
    "schluessel","straftat","kreisart","faelle","hz_pro_100k",
    "versuche_anzahl","versuche_prozent",
    "mit_schusswaffe_gedroht","mit_schusswaffe_geschossen",
    "aufklaerung_faelle","aq_prozent",
    "tatverdaechtige_insgesamt","tatverdaechtige_maennlich","tatverdaechtige_weiblich",
    "nichtdeutsche_tv_anzahl","nichtdeutsche_tv_anteil_prozent",
    "jahr",
]

def to_number(x):
    """Text -> Zahl (Tausender raus, Komma -> Punkt)."""
    if pd.isna(x):
        return x
    s = str(x).strip()
    if s == "":
        return pd.NA
    s = s.replace("\u202f", "").replace("\u00a0", "").replace(" ", "")
    s = s.replace(",", ".").replace("%", "")
    try:
        if re.match(r"^-?\d+\.\d+$", s):
            return float(s)
        if re.match(r"^-?\d+$", s):
            return int(s)
        return float(s)
    except Exception:
        return pd.NA

def process_file(xlsx_path: Path) -> pd.DataFrame:
    """Ein Jahr verarbeiten und DataFrame zurückgeben."""
    print(f"\n[DEBUG] === Datei starten: {xlsx_path.name} ===")

    # Header finden
    header_top = find_header_row(xlsx_path, SHEET_NAME)
    header_bottom = header_top + 1
    print(f"[DEBUG] Header-Zeilen: top={header_top}, bottom={header_bottom}")

    # Daten laden (2 Level Header)
    df = pd.read_excel(xlsx_path, sheet_name=SHEET_NAME, header=[header_top, header_bottom])

    # Header reparieren (unten füllen)
    group_counters = {k: 0 for k in GROUP_SPECS.keys()}
    fixed_cols = []
    for top, bottom in df.columns.to_list():
        top_s = "" if top is None else str(top).strip()
        bot_s = "" if bottom is None else str(bottom).strip()
        if top_s in GROUP_SPECS and is_unnamed(bot_s):
            idx = group_counters[top_s]
            if idx < len(GROUP_SPECS[top_s]):
                bot_s = GROUP_SPECS[top_s][idx]
                group_counters[top_s] += 1
            else:
                bot_s = ""
        if is_unnamed(bot_s):
            bot_s = ""
        fixed_cols.append((top_s, bot_s))
    df.columns = pd.MultiIndex.from_tuples(fixed_cols)

    # Spalten lexikografisch sortieren (stabiler Zugriff)
    df = df.sort_index(axis=1)
    print("[DEBUG] Columns lexsorted.")

    # Helper: sichere Series holen
    def col_series(frame: pd.DataFrame, top: str, bottom: str = "") -> pd.Series:
        """Hole eine Spalte als Serie."""
        sel = frame.loc[:, (top, bottom)]
        if isinstance(sel, pd.DataFrame):
            if sel.shape[1] != 1:
                print("[DEBUG] Warn: Mehrere Spalten für", (top, bottom), "-> erste genommen.")
            return sel.iloc[:, 0]
        return sel

    # Filter: nur Wetteraukreis
    if ("Stadt-/Landkreis", "") not in df.columns:
        raise ValueError("Spalte 'Stadt-/Landkreis' fehlt.")
    kreis_col = col_series(df, "Stadt-/Landkreis", "")
    df = df[kreis_col.astype(str).str.strip() == "Wetteraukreis"]
    print("[DEBUG] Zeilen nach Filter Wetteraukreis:", len(df))

    # Auswahl + Umbenennen
    rename_map = {raw: new for raw, new in RAW_TO_FINAL.items() if raw in df.columns}
    df_sel = df[list(rename_map.keys())].copy()
    df_sel = df_sel.sort_index(axis=1)  # stabil

    # HZ-Duplikate (falls beide)
    if ("HZ", "") in df_sel.columns and ("HZ nach Zensus 2022", "") in df_sel.columns:
        df_sel = df_sel.drop(columns=[("HZ nach Zensus 2022", "")])

    # Umbenennen Level 0, dann flach machen
    df_sel = df_sel.rename(columns=rename_map, level=0)
    df_sel.columns = [RAW_TO_FINAL[(c if isinstance(c, tuple) else (c, ""))] for c in df_sel.columns.to_flat_index()]

    # HZ sicherstellen (falls aus anderer Quelle)
    if "hz_pro_100k" not in df_sel.columns:
        if ("HZ", "") in df.columns:
            df_sel["hz_pro_100k"] = col_series(df, "HZ", "")
        elif ("HZ nach Zensus 2022", "") in df.columns:
            df_sel["hz_pro_100k"] = col_series(df, "HZ nach Zensus 2022", "")
        else:
            raise ValueError("HZ-Spalte nicht gefunden.")

    # Zahlen normalisieren
    num_cols = [
        "faelle","hz_pro_100k","versuche_anzahl","versuche_prozent",
        "mit_schusswaffe_gedroht","mit_schusswaffe_geschossen",
        "aufklaerung_faelle","aq_prozent",
        "tatverdaechtige_insgesamt","tatverdaechtige_maennlich","tatverdaechtige_weiblich",
        "nichtdeutsche_tv_anzahl","nichtdeutsche_tv_anteil_prozent",
    ]
    for c in [c for c in num_cols if c in df_sel.columns]:
        df_sel[c] = df_sel[c].apply(to_number)

    # jahr aus Dateiname
    m = re.search(r"kriminalstatistik_(\d{4})\.xlsx$", xlsx_path.name)
    jahr = int(m.group(1)) if m else None
    df_sel["jahr"] = jahr

    # Debug
    print("[DEBUG] Spalten:", list(df_sel.columns))
    print("[DEBUG] Kopf:")
    print(df_sel.head(3))

    return df_sel

def export_datenkatalog(path: Path):
    """Schreibe Mapping (Original -> neu) als CSV."""
    rows = []
    for (top, bottom), final in RAW_TO_FINAL.items():
        orig = top if (bottom == "" or bottom is None) else f"{top} | {bottom}"
        rows.append({"original_header": orig, "final_name": final})
    df_dict = pd.DataFrame(rows).drop_duplicates().sort_values("final_name")
    df_dict.to_csv(path, sep=";", index=False, encoding="utf-8")
    print(f"[DEBUG] Datenkatalog gespeichert: {path}")

def main():
    """Alle Dateien verarbeiten und CSV+Datenkatalog schreiben."""
    files = sorted(BASE_DIR.glob("kriminalstatistik_*.xlsx"))
    print("[DEBUG] Gefundene Dateien:", [f.name for f in files])
    if not files:
        raise SystemExit("Keine Dateien gefunden.")

    frames = []
    for p in files:
        frames.append(process_file(p))

    df_all = pd.concat(frames, ignore_index=True)

    # Reihenfolge sichern
    missing_final = [c for c in FINAL_ORDER if c not in df_all.columns]
    if missing_final:
        raise ValueError(f"Fehlende Zielspalten: {missing_final}")
    df_all = df_all[FINAL_ORDER]

    # Leichte Checks
    def in_0_100(series):
        return series.dropna().between(0, 100).all()

    print("\n[DEBUG] Gesamt-Zeilen:", len(df_all))
    print("[DEBUG] Jahre vorhanden:", sorted(df_all['jahr'].unique()))
    if "versuche_anzahl" in df_all and "faelle" in df_all:
        print("[DEBUG] Check versuche_anzahl <= faelle:",
              (df_all["versuche_anzahl"].dropna() <= df_all["faelle"].dropna()).all())
    for pc in ["versuche_prozent","aq_prozent","nichtdeutsche_tv_anteil_prozent"]:
        print(f"[DEBUG] Check 0..100 für {pc}:", in_0_100(df_all[pc]))

    # CSV schreiben
    df_all.to_csv(OUT_CSV, sep=";", index=False, encoding="utf-8")
    print(f"[DEBUG] CSV gespeichert: {OUT_CSV}")

    # Datenkatalog schreiben (Header-Übersetzung)
    export_datenkatalog(OUT_DICT)

if __name__ == "__main__":
    main()
