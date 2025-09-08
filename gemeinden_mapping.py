import csv

gemeinden = [
    (6, "Land Hessen", "06000000"),
    (6440, "Wetteraukreis", "06440000"),
    (6440001, "Altenstadt", "06440001"),
    (6440002, "Bad Nauheim", "06440002"),
    (6440003, "Bad Vilbel", "06440003"),
    (6440004, "Büdingen", "06440004"),
    (6440005, "Butzbach", "06440005"),
    (6440006, "Echzell", "06440006"),
    (6440007, "Florstadt", "06440007"),
    (6440008, "Friedberg (Hessen)", "06440008"),
    (6440009, "Gedern", "06440009"),
    (6440010, "Glauburg", "06440010"),
    (6440011, "Hirzenhain", "06440011"),
    (6440012, "Karben", "06440012"),
    (6440013, "Kefenrod", "06440013"),
    (6440014, "Limeshain", "06440014"),
    (6440015, "Münzenberg", "06440015"),
    (6440016, "Nidda", "06440016"),
    (6440017, "Niddatal", "06440017"),
    (6440018, "Ober-Mörlen", "06440018"),
    (6440019, "Ortenberg", "06440019"),
    (6440020, "Ranstadt", "06440020"),
    (6440021, "Reichelsheim (Wetterau)", "06440021"),
    (6440022, "Rockenberg", "06440022"),
    (6440023, "Rosbach v. d. Höhe", "06440023"),
    (6440024, "Wölfersheim", "06440024"),
    (6440025, "Wöllstadt", "06440025"),
]

OUTPUT_FILENAME = "result/gemeinden_mapping_tabelle.csv"

def main():
    with open(OUTPUT_FILENAME, mode="w", newline="", encoding="utf-8") as csvfile:
        writer = csv.writer(csvfile, delimiter=";")
        writer.writerow(["gemeinde_id", "gemeinde", "gemeinde_schluessel"])
        writer.writerows(gemeinden)

    print(f"Result saved to {OUTPUT_FILENAME}")

if __name__ == "__main__":
    main()
