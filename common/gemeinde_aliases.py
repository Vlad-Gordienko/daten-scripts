# Dictionary of canonical Gemeinde (municipality) names and their known aliases or name variants.
# Used for normalizing Gemeinde names extracted from filenames or datasets where formatting may vary.
# Ensures consistent naming across all data sources and output files.

# Example:
# - "Friedberg Hessen" and "Friedberg" will both be mapped to "Friedberg (Hessen)"
# - "Reichelsheim Wetterau" → "Reichelsheim (Wetterau)"

gemeinde_aliases = {
    "Altenstadt": ["Altenstadt"],
    "Bad Nauheim": ["Bad Nauheim", "Bad Nauheim, Stadt", "Bad Nauheim Stadt"],
    "Bad Vilbel": ["Bad Vilbel", "Bad Vilbel, Stadt", "Bad Vilbel Stadt"],
    "Büdingen": ["Büdingen", "Büdingen, Stadt", "Büdingen Stadt"],
    "Butzbach": [
        "Butzbach",
        "Butzbach Friedrich-Ludwig-Weidig-Stadt",
        "Butzbach, Fried.-L.-Weidig-St.",
        "Butzbach, Friedrich-Ludwig-Weidig-Stadt",
    ],
    "Echzell": ["Echzell"],
    "Florstadt": ["Florstadt", "Florstadt, Stadt", "Florstadt Stadt"],
    "Friedberg (Hessen)": [
        "Friedberg (Hessen)",
        "Friedberg",
        "Friedberg Hessen",
        "Friedberg Hessen Stadt",
        "Friedberg (Hessen), Kreisstadt",
        "Friedberg (Hessen), Stadt",
    ],
    "Gedern": ["Gedern", "Gedern, Stadt", "Gedern Stadt"],
    "Glauburg": ["Glauburg"],
    "Hirzenhain": ["Hirzenhain"],
    "Karben": ["Karben", "Karben, Stadt", "Karben Stadt"],
    "Kefenrod": ["Kefenrod"],
    "Limeshain": ["Limeshain"],
    "Münzenberg": ["Münzenberg", "Münzenberg, Stadt", "Münzenberg Stadt"],
    "Nidda": ["Nidda", "Nidda, Stadt", "Nidda Stadt"],
    "Niddatal": ["Niddatal", "Niddatal, Stadt", "Niddatal Stadt"],
    "Ober-Mörlen": ["Ober-Mörlen"],
    "Ortenberg": ["Ortenberg", "Ortenberg, Stadt", "Ortenberg Stadt"],
    "Ranstadt": ["Ranstadt"],
    "Reichelsheim (Wetterau)": [
        "Reichelsheim (Wetterau)",
        "Reichelsheim",
        "Reichelsheim Wetterau",
        "Reichelsheim Wetterau Stadt",
        "Reichelsheim (Wetterau), Stadt",
    ],
    "Rockenberg": ["Rockenberg"],
    "Rosbach v. d. Höhe": [
        "Rosbach v. d. Höhe",
        "Rosbach v d Höhe",
        "Rosbach v d Höhe Stadt",
        "Rosbach v. d. Höhe, Stadt",
    ],
    "Wölfersheim": ["Wölfersheim"],
    "Wöllstadt": ["Wöllstadt"],
    "Wetteraukreis": ["Wetteraukreis"],
}
