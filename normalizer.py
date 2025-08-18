import os
import re
import logging
import pandas as pd

# === Константы ===
INPUT_DIR = "data"
OUTPUT_DIR = "result"
FILENAME = "Bevölkerungsbewegung.xlsx"   # можешь заменить на любой входной файл
SHEET_NAME = None                        # если None — берётся первый лист

INPUT_PATH = os.path.join(INPUT_DIR, FILENAME)
BASENAME, EXT = os.path.splitext(FILENAME)
OUTPUT_PATH = os.path.join(OUTPUT_DIR, f"{BASENAME}_normalized.xlsx")
LOG_PATH = os.path.join(OUTPUT_DIR, f"{BASENAME}_normalized.log")

# === Импорты твоего маппера ===
# Ориентируюсь на твой пример: функции доступны из common.mapping
from common.mapping import (
    get_gemeinde_from_gebiet,
    normalize_gemeinde_name,
    track_undetected_gebiete,
    log_missing_gebiete,
)

# === Логирование ===
os.makedirs(OUTPUT_DIR, exist_ok=True)
logging.basicConfig(
    level=logging.INFO,
    format="%(levelname)s: %(message)s",
    handlers=[logging.StreamHandler(), logging.FileHandler(LOG_PATH, mode="w", encoding="utf-8")],
)

# === Алиасы колонок (регистронезависимо) ===
YEAR_ALIASES = ("jahr", "year", "jahrgang")
GEMEINDE_ALIASES = ("gemeinde",)
GEBIET_ALIASES = ("gebiet",)

# Регекс «числоподобной» строки: опциональный минус (любой вид), цифры и пробелы-разделители
NEG_CHARS = "\u2212\u2012\u2013\u2014\u2015-"  # − ‒ – — ― -
NUM_LIKE_RE = re.compile(rf"^[\s{NEG_CHARS}]?\d[\d\s]*$")


def find_column(df: pd.DataFrame, aliases: tuple[str, ...]) -> str | None:
    """Находит имя колонки по алиасам (регистронезависимо). Возвращает реальное имя колонки или None."""
    lower_map = {c.lower(): c for c in df.columns}
    for a in aliases:
        if a in lower_map:
            return lower_map[a]
    return None


def to_int_if_number_like(val):
    """
    Приводит строку вида "1 234" / "–7" / "– 46" к int.
    """
    if pd.isna(val):
        return val
    if isinstance(val, (int, float)) and not isinstance(val, bool):
        if float(val).is_integer():
            return int(val)
        return val

    s = str(val).strip()
    if not s:
        return val

    # заменяем все варианты минуса на "-"
    for ch in NEG_CHARS:
        s = s.replace(ch, "-")

    # убираем пробелы
    s = s.replace(" ", "")

    # кейс вида "-46" или "1234"
    if s.lstrip("-").isdigit():
        try:
            return int(s)
        except Exception:
            logging.debug(f"Cannot cast to int: raw='{val}' cleaned='{s}'")
            return val

    return val


def normalize_numeric_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Применяет to_int_if_number_like ко всем object-колонкам, кроме явных текстовых идентификаторов.
    Не трогаем столбцы 'gemeinde'/'gebiet' (любой регистр).
    """
    skip = {c for c in df.columns if c.lower() in (*GEMEINDE_ALIASES, *GEBIET_ALIASES)}
    for col in df.columns:
        if col in skip:
            continue
        if df[col].dtype == "object":
            df[col] = df[col].map(to_int_if_number_like)
    return df


def ensure_year_column(df: pd.DataFrame) -> tuple[pd.DataFrame, str]:
    """
    Находит/переименовывает годовую колонку в 'jahr' и приводит к int.
    Отбрасывает строки без валидного года.
    """
    year_col = find_column(df, YEAR_ALIASES)
    if not year_col:
        raise ValueError("Не найдена колонка года (ожидается одна из: jahr | year | jahrgang)")
    df = df.rename(columns={year_col: "jahr"})
    # Приводим к int через to_int_if_number_like и выбрасываем NaN/неинты
    df["jahr"] = df["jahr"].map(to_int_if_number_like)
    df = df[pd.to_numeric(df["jahr"], errors="coerce").notna()].copy()
    df["jahr"] = df["jahr"].astype(int)
    return df, "jahr"


def ensure_gemeinde_column(df: pd.DataFrame) -> tuple[pd.DataFrame, str]:
    """
    Гарантирует наличие колонки 'gemeinde'.
    - если есть 'gemeinde' → нормализация по normalize_gemeinde_name
    - иначе если есть 'gebiet' → маппинг через get_gemeinde_from_gebiet
        * пустая строка (ignore) → дроп строки
        * строка/класс → переносим в 'gemeinde' как есть
    - если нет обеих → дроп строки
    """
    g_col = find_column(df, GEMEINDE_ALIASES)
    gb_col = find_column(df, GEBIET_ALIASES)

    if g_col:
        df = df.rename(columns={g_col: "gemeinde"})
        # нормализация имён по алиасам
        df["gemeinde"] = df["gemeinde"].astype(str).map(normalize_gemeinde_name)
        # пустые после нормализации — дроп
        df = df[df["gemeinde"].astype(str).str.strip() != ""].copy()
        return df, "gemeinde"

    if gb_col:
        df = df.rename(columns={gb_col: "gebiet"})
        # лог полноты покрытия
        all_gebieten = df["gebiet"].dropna().astype(str).unique().tolist()
        found = track_undetected_gebiete(all_gebieten)
        log_missing_gebiete(found)

        def map_gebiet(val):
            if pd.isna(val):
                return ""
            res = get_gemeinde_from_gebiet(str(val).strip())
            if res == "":
                return ""  # игнор-лист — дропнем ниже
            return res  # это может быть строка или класс — оставляем как есть

        df["gemeinde"] = df["gebiet"].map(map_gebiet)
        before = len(df)
        df = df[df["gemeinde"].astype(str).str.strip() != ""].copy()
        dropped = before - len(df)
        if dropped:
            logging.info(f"Удалено строк по ignore_list или пустым Gemeinde: {dropped}")
        return df, "gemeinde"

    # нет ни gemeinde, ни gebiet — всё выкидываем
    logging.warning("Не найдены колонки 'gemeinde' или 'gebiet' — все строки будут отброшены.")
    return df.iloc[0:0].copy(), "gemeinde"


def sort_by_gemeinde_and_year(df: pd.DataFrame) -> pd.DataFrame:
    """
    Группирует строки по Gemeinde и сортирует годы по возрастанию.
    Дополнительно сортирует по "уровню" ключа: Land (len=1) → Kreis (len=4) → Gemeinde (len=7).
    """
    # приведём год к int на всякий случай
    df["jahr"] = pd.to_numeric(df["jahr"], errors="coerce").astype("Int64")
    df = df[df["jahr"].notna()].copy()
    df["jahr"] = df["jahr"].astype(int)

    # вычисляем уровень по длине gemeinde_schluessel, если колонка есть
    if "gemeinde_schluessel" in df.columns:
        lvl = df["gemeinde_schluessel"].astype(str).str.len()
    else:
        lvl = 99  # если нет ключа — отправим в конец

    df = df.assign(
        _lvl=lvl,
        _g=df["gemeinde"].astype(str),
    )

    df = df.sort_values(by=["_lvl", "_g", "jahr"], ascending=[True, True, True]) \
           .drop(columns=["_lvl", "_g"])

    return df


def parse_excel():
    """
    Main entry point:
    - Load Excel (sheet or first sheet)
    - Ensure 'jahr' and 'gemeinde'
    - Normalize numeric-like strings to int
    - Sort by 'jahr' (asc) then 'gemeinde' (asc)
    - Save to result/<basename>_normalized.xlsx
    """
    if not os.path.exists(INPUT_PATH):
        raise FileNotFoundError(f"Входной файл не найден: {INPUT_PATH}")

    logging.info(f"Чтение: {INPUT_PATH}")
    xls = pd.ExcelFile(INPUT_PATH)
    sheet = SHEET_NAME if SHEET_NAME is not None else xls.sheet_names[0]
    logging.info(f"Лист: {sheet}")

    df = pd.read_excel(xls, sheet_name=sheet, dtype=object)

    # Обязательные ключи
    df, year_col = ensure_year_column(df)
    df, gemeinde_col = ensure_gemeinde_column(df)

    # Нормализация чисел (все object-колонки, кроме идентификаторов)
    df = normalize_numeric_columns(df)

    # Сортировка
    df = sort_by_gemeinde_and_year(df)

    # Запись
    logging.info(f"Сохранение результата: {OUTPUT_PATH}")
    df.to_excel(OUTPUT_PATH, index=False)
    logging.info("Готово.")


if __name__ == "__main__":
    parse_excel()
