import os
import pandas as pd

# === Константы ===
OUTPUT_DIR = "result"
FILENAME = "gemeinden_unique.xlsx"
OUTPUT_PATH = os.path.join(OUTPUT_DIR, FILENAME)

# === Данные (в порядке первого появления) ===
DATA = [
    ("6", "Land Hessen"),
    ("6440", "Wetteraukreis"),
    ("6440001", "Altenstadt"),
    ("6440002", "Bad Nauheim"),
    ("6440003", "Bad Vilbel"),
    ("6440005", "Butzbach"),
    ("6440004", "Büdingen"),
    ("6440006", "Echzell"),
    ("6440007", "Florstadt"),
    ("6440008", "Friedberg (Hessen)"),
    ("6440009", "Gedern"),
    ("6440010", "Glauburg"),
    ("6440011", "Hirzenhain"),
    ("6440011", "ixbt"),
    ("6440012", "Karben"),
    ("6440013", "Kefenrod"),
    ("6440014", "Limeshain"),
    ("6440015", "Münzenberg"),
    ("6440016", "Nidda"),
    ("6440017", "Niddatal"),
    ("6440018", "Ober-Mörlen"),
    ("6440019", "Ortenberg"),
    ("6440020", "Ranstadt"),
    ("6440021", "Reichelsheim (Wetterau)"),
    ("6440022", "Rockenberg"),
    ("6440023", "Rosbach v. d. Höhe"),
    ("6440024", "Wölfersheim"),
    ("6440025", "Wöllstadt"),
]

def generate_excel():
    """
    Создаёт Excel-файл с колонками:
    - gemeinde_schluessel (строка)
    - gemeinde (строка)
    """
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    df = pd.DataFrame(DATA, columns=["gemeinde_schluessel", "gemeinde"])
    # сохраняем ключи как строки (без научной записи и потерь)
    df["gemeinde_schluessel"] = df["gemeinde_schluessel"].astype(str)

    # Если понадобится zero-pad до 8 знаков — раскомментируй:
    # df["gemeinde_schluessel"] = df["gemeinde_schluessel"].str.zfill(8)

    df.to_excel(OUTPUT_PATH, index=False)
    print(f"Saved: {OUTPUT_PATH}")

if __name__ == "__main__":
    generate_excel()
