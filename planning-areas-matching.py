#!/usr/bin/env python3

import pandas as pd
from pathlib import Path


INPUT_FILE = Path("data/WK_Planungsraeume.xlsx")
OUTPUT_FILE = Path("WK_Planungsraeume_matching.csv")


def find_header_row(df_no_header):
    """Ищем строку с заголовками."""
    must_have = {"PLZ", "Gemeindeziffer", "Gemeinde", "ASD-Regionen"}
    for i in range(min(50, len(df_no_header))):
        vals = [str(x).strip() for x in df_no_header.iloc[i].tolist()]
        if must_have.issubset(set(vals)):
            return i
    with pd.option_context("display.max_columns", None, "display.width", 200):
        print(df_no_header.head(5))
    raise ValueError("Не удалось найти строку с заголовками (PLZ/Gemeindeziffer/Gemeinde/ASD-Regionen).")


def read_zustaendigkeiten(xlsx_path: Path) -> pd.DataFrame:
    """Читаем лист 'Zuständigkeiten'."""
    raw = pd.read_excel(xlsx_path, sheet_name="Zuständigkeiten", header=None, dtype=str)

    header_row = find_header_row(raw)
    df = pd.read_excel(xlsx_path, sheet_name="Zuständigkeiten", header=header_row, dtype=str)
    return df


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Оставляем нужные колонки и переименовываем."""
    orig_cols = list(df.columns)
    df.columns = [str(c).strip() for c in df.columns]

    needed_raw = ["ASD-Regionen", "Gemeinde", "Gemeindeziffer", "PLZ"]
    missing = [c for c in needed_raw if c not in df.columns]
    if missing:
        raise ValueError(f"Не найдены обязательные колонки: {missing}")

    ren = {
        "ASD-Regionen": "ASD-Region",
        "Gemeindeziffer": "Gemeindekennziffer",
        "Gemeinde": "Gemeinde",
        "PLZ": "PLZ",
    }

    return df[needed_raw].rename(columns=ren)


def clean_and_format(df: pd.DataFrame) -> pd.DataFrame:
    for col in ["ASD-Region", "Gemeinde", "Gemeindekennziffer", "PLZ"]:
        df[col] = df[col].astype(str).str.strip()

    df["PLZ"] = (
        df["PLZ"]
        .str.extract(r"(\d+)", expand=False)
        .fillna("")
        .str.zfill(5)
    )
    df["Gemeindekennziffer"] = (
        df["Gemeindekennziffer"]
        .str.extract(r"(\d+)", expand=False)
        .fillna("")
        .str.zfill(8)
    )

    before = len(df)
    df = df[(df["Gemeinde"] != "") & (df["PLZ"] != "") & (df["Gemeinde"] != "") & (df["Gemeindekennziffer"] != "")]

    df = df.drop_duplicates(subset=["Gemeindekennziffer", "PLZ", "ASD-Region", "Gemeinde"])
    df["GEOjson"] = ""

    df = df[["ASD-Region", "Gemeinde", "Gemeindekennziffer", "PLZ", "GEOjson"]].reset_index(drop=True)

    return df


def build_matching_table(xlsx_path: Path) -> pd.DataFrame:
    xls = pd.ExcelFile(xlsx_path)

    df_raw = read_zustaendigkeiten(xlsx_path)
    df_norm = normalize_columns(df_raw)
    df_clean = clean_and_format(df_norm)
    return df_clean


def main():
    if not INPUT_FILE.exists():
        print(f"Ошибка: файл не найден {INPUT_FILE}")
        return

    try:
        df = build_matching_table(INPUT_FILE)
    except Exception as e:
        print("Ошибка при создании таблицы:", repr(e))
        return

    try:
        df.to_csv(OUTPUT_FILE, index=False, encoding="utf-8-sig")
        print(f"Готово: записано в {OUTPUT_FILE}")
    except Exception as e:
        print("Ошибка при сохранении CSV:", repr(e))


if __name__ == "__main__":
    main()
