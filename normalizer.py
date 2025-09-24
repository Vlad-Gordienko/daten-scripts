#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
normalizer.py — sehr einfacher CSV-Fixer.
- fester Trenner ist ';', kein Auto-Detekt
- einfache Zeichen-Ersetzung mit Map
- gleiche Anzahl Spalten wie Kopfzeile
- Datei speichern neben Skript mit gleichem Namen (überschreiben)
"""

import argparse
import csv
import io
import os
import re
import sys
import logging
from collections import Counter
from typing import List

# --- kleine Zeichen-Liste zum Ersetzen (kann größer werden) ---
CHAR_MAP = {
    "Ä": "AE", "Ö": "OE", "Ü": "UE",
    "ä": "ae", "ö": "oe", "ü": "ue",
    "ß": "SS",
    "€": "EUR",
    "\u00A0": " ",     # NBSP → Leerzeichen
    "“": '"', "”": '"', "„": '"', "«": '"', "»": '"',
    "’": "'", "‚": "'",
    "—": "-", "–": "-",
    "\u200B": "", "\u200C": "", "\u200D": "", "\uFEFF": "",  # zero-width, BOM
}

# löschen von Steuer-Zeichen (außer \t,\n,\r, diese macht csv selbst)
_CTRL_RE = re.compile(r"[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F\u0080-\u009F]")


def normalize_text(s: str, unknown_counter: Counter) -> str:
    """Einfache Normalisierung einer Zelle."""
    if not s:
        return s

    # 1) Ersetzen mit Map
    out = []
    for ch in s:
        out.append(CHAR_MAP.get(ch, ch))
    s = "".join(out)

    # 2) Steuer-Zeichen löschen
    s = _CTRL_RE.sub("", s)

    # 3) unbekannte Zeichen merken (nicht-ASCII)
    for ch in s:
        o = ord(ch)
        if o in (9, 10, 13):  # Tab / neue Zeile
            continue
        if o < 32 or o > 126:
            unknown_counter[ch] += 1

    return s


def process_one_file(input_path: str, enc: str, delimiter: str) -> None:
    """Liest Datei und schreibt neue Datei neben dem Skript (gleicher Dateiname)."""
    # Ausgabe-Pfad: Ordner vom Skript
    script_dir = os.path.dirname(os.path.abspath(__file__))
    output_path = os.path.join(script_dir, os.path.basename(input_path))

    unknown = Counter()
    truncated = padded = 0
    total_rows = 0

    # Eingabe lesen
    with io.open(input_path, "r", encoding=enc, newline="") as rf:
        reader = csv.reader(rf, delimiter=delimiter, quotechar='"', doublequote=True)
        # Kopfzeile lesen
        header: List[str] = next(reader, [])
        header = [normalize_text(x, unknown) for x in header]
        expected_cols = len(header) if header else 0

        # wenn Kopf leer ist → erste Datenzeile nehmen für Länge
        first_row = None
        if expected_cols == 0:
            first_row = next(reader, [])
            first_row = [normalize_text(x, unknown) for x in first_row]
            expected_cols = max(1, len(first_row))
            header = [f"col_{i}" for i in range(1, expected_cols + 1)]

        # Ausgabe schreiben
        with io.open(output_path, "w", encoding="utf-8", newline="") as wf:
            writer = csv.writer(
                wf,
                delimiter=delimiter,
                quotechar='"',
                doublequote=True,
                lineterminator="\n",
                quoting=csv.QUOTE_MINIMAL,
            )

            # Kopf an expected_cols anpassen
            if len(header) > expected_cols:
                header = header[:expected_cols]
            elif len(header) < expected_cols:
                header += [""] * (expected_cols - len(header))
            writer.writerow(header)

            # wenn erste Datenzeile Kopf-Ersatz war → jetzt schreiben
            if first_row is not None:
                total_rows += 1
                row = first_row
                if len(row) > expected_cols:
                    row = row[:expected_cols]
                    truncated += 1
                elif len(row) < expected_cols:
                    row += [""] * (expected_cols - len(row))
                    padded += 1
                writer.writerow(row)

            # alle anderen Zeilen
            for row in reader:
                total_rows += 1
                row = [normalize_text(x, unknown) for x in row]

                # 1) leere Zellen am Ende löschen (wegen ;;;)
                while row and (row[-1] is None or row[-1] == ""):
                    row.pop()

                # 2) Länge an expected_cols anpassen
                if len(row) > expected_cols:
                    row = row[:expected_cols]
                    truncated += 1
                elif len(row) < expected_cols:
                    row += [""] * (expected_cols - len(row))
                    padded += 1
                writer.writerow(row)

    # kurzer Bericht auf stderr
    sys.stderr.write(
        f"[ok] {os.path.basename(input_path)} → {os.path.basename(output_path)} | "
        f"rows={total_rows}, truncated={truncated}, padded={padded}\n"
    )
    if unknown:
        sys.stderr.write("    [warn] unbekannte nicht-ASCII Zeichen:\n")
        for ch, cnt in unknown.most_common(20):
            sys.stderr.write(f"        {repr(ch)} (U+{ord(ch):04X}) × {cnt}\n")


def main() -> int:
    parser = argparse.ArgumentParser(description="Sehr einfacher CSV-Normalizer.")
    parser.add_argument("path", help="Pfad zu CSV-Datei")
    parser.add_argument(
        "--encoding", default="cp1252", help="Eingabe-Encoding (Standard: cp1252)"
    )
    parser.add_argument(
        "--delimiter", default=";", help="CSV-Trenner (Standard: ';')"
    )
    args = parser.parse_args()

    # Logging kurz initialisieren (geht trotzdem alles auf stderr)
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

    if not os.path.exists(args.path):
        sys.stderr.write(f"[err] Datei nicht gefunden: {args.path}\n")
        return 1

    try:
        process_one_file(args.path, args.encoding, args.delimiter)
        logging.info("Fertig.")
        return 0
    except UnicodeDecodeError as e:
        sys.stderr.write(
            f"[err] Lese-Fehler im Encoding '{args.encoding}': {e}\n"
        )
        return 2
    except csv.Error as e:
        sys.stderr.write(f"[err] CSV-Fehler: {e}\n")
        return 3
    except Exception as e:
        sys.stderr.write(f"[err] Unerwarteter Fehler: {e}\n")
        return 99


if __name__ == "__main__":
    sys.exit(main())
