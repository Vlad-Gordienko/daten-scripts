#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
import csv
import os

CHAR_MAP = {
    "Ä": "AE", "Ö": "OE", "Ü": "UE",
    "ä": "ae", "ö": "oe", "ü": "ue",
    "ß": "SS",
    "€": "EUR", "$": "USD"
}

def normalize_value(value):
    if value is None:
        return ""
    value = value.strip()
    if not value:
        return ""
    result = []
    for ch in value:
        if ch in CHAR_MAP:
            result.append(CHAR_MAP[ch])
        elif ord(ch) < 128:
            result.append(ch)
        else:
            result.append("*")
    return "".join(result)

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("usage: python3 normalizer.py <input.csv>")
        sys.exit(1)

    input_path = sys.argv[1]
    base_name = os.path.basename(input_path)
    output_path = os.path.join(os.path.dirname(__file__), base_name)

    with open(input_path, "r", encoding="utf-8", errors="replace") as infile, \
         open(output_path, "w", newline='', encoding="utf-8") as outfile:
        reader = csv.reader(infile, delimiter=';')
        writer = csv.writer(outfile, delimiter=';')
        for row in reader:
            writer.writerow([normalize_value(v) for v in row])
