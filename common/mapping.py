import sys
import os
sys.path.append(os.path.dirname(__file__))

import logging

from gemeinden.altenstadt import mapping as altenstadt_mapping
from gemeinden.bad_nauheim import mapping as bad_nauheim_mapping
from gemeinden.bad_vilbel import mapping as bad_vilbel_mapping
from gemeinden.buedingen import mapping as buedingen_mapping
from gemeinden.butzbach import mapping as butzbach_mapping
from gemeinden.echzell import mapping as echzell_mapping
from gemeinden.florstadt import mapping as florstadt_mapping
from gemeinden.friedberg_hessen import mapping as friedberg_hessen_mapping
from gemeinden.gedern import mapping as gedern_mapping
from gemeinden.glauburg import mapping as glauburg_mapping
from gemeinden.hirzenhain import mapping as hirzenhain_mapping
from gemeinden.karben import mapping as karben_mapping
from gemeinden.kefenrod import mapping as kefenrod_mapping
from gemeinden.limeshain import mapping as limeshain_mapping
from gemeinden.muenzenberg import mapping as muenzenberg_mapping
from gemeinden.nidda import mapping as nidda_mapping
from gemeinden.niddatal import mapping as niddatal_mapping
from gemeinden.ober_moerlen import mapping as ober_moerlen_mapping
from gemeinden.ortenberg import mapping as ortenberg_mapping
from gemeinden.ranstadt import mapping as ranstadt_mapping
from gemeinden.reichelsheim_wetterau import mapping as reichelsheim_wetterau_mapping
from gemeinden.rockenberg import mapping as rockenberg_mapping
from gemeinden.rosbach_v_d_hoehe import mapping as rosbach_v_d_hoehe_mapping
from gemeinden.woelfersheim import mapping as woelfersheim_mapping
from gemeinden.woellstadt import mapping as woellstadt_mapping

from schluessel_map import schluessel_map
from gemeinde_aliases import gemeinde_aliases

from gemeinden.gemeinden import (
    Altenstadt,
    Bad_Nauheim,
    Bad_Vilbel,
    Buedingen,
    Butzbach,
    Echzell,
    Florstadt,
    Friedberg_Hessen,
    Gedern,
    Glauburg,
    Hirzenhain,
    Karben,
    Kefenrod,
    Limeshain,
    Muenzenberg,
    Nidda,
    Niddatal,
    Ober_Moerlen,
    Ortenberg,
    Ranstadt,
    Reichelsheim_Wetterau,
    Rockenberg,
    Rosbach_v_d_Hoehe,
    Woelfersheim,
    Woellstadt,
)

# Imports all municipality-to-gebiet mappings from individual submodules.
# Each imported mapping contains lists of alternative Gebiet names grouped under one Gemeinde.

mapping = {
    # Flatten all imported mappings into one combined dict
    **altenstadt_mapping,
    **bad_nauheim_mapping,
    **bad_vilbel_mapping,
    **buedingen_mapping,
    **butzbach_mapping,
    **echzell_mapping,
    **florstadt_mapping,
    **friedberg_hessen_mapping,
    **gedern_mapping,
    **glauburg_mapping,
    **hirzenhain_mapping,
    **karben_mapping,
    **kefenrod_mapping,
    **limeshain_mapping,
    **muenzenberg_mapping,
    **nidda_mapping,
    **niddatal_mapping,
    **ober_moerlen_mapping,
    **ortenberg_mapping,
    **ranstadt_mapping,
    **reichelsheim_wetterau_mapping,
    **rockenberg_mapping,
    **rosbach_v_d_hoehe_mapping,
    **woelfersheim_mapping,
    **woellstadt_mapping,
}

# Dictionary defining canonical Gemeinde names and their alternative spellings or representations
mapping_gemeinde = {
    Altenstadt: ['Altenstadt'],
    Bad_Nauheim: ['Bad Nauheim'],
    Bad_Vilbel: ['Bad Vilbel'],
    Buedingen: ['Büdingen'],
    Butzbach: ['Butzbach'],
    Echzell: ['Echzell'],
    Florstadt: ['Florstadt'],
    Friedberg_Hessen: ['Friedberg (Hessen)', 'Frieberg'],
    Gedern: ['Gedern'],
    Glauburg: ['Glauburg'],
    Hirzenhain: ['Hirzenhain'],
    Karben: ['Karben'],
    Kefenrod: ['Kefenrod'],
    Limeshain: ['Limeshain'],
    Muenzenberg: ['Münzenberg'],
    Nidda: ['Nidda'],
    Niddatal: ['Niddatal'],
    Ober_Moerlen: ['Ober-Mörlen'],
    Ortenberg: ['Ortenberg'],
    Ranstadt: ['Ranstadt'],
    Reichelsheim_Wetterau: ['Reichelsheim (Wetterau)'],
    Rockenberg: ['Rockenberg'],
    Rosbach_v_d_Hoehe: ['Rosbach v. d. Höhe'],
    Woelfersheim: ['Wölfersheim'],
    Woellstadt: ['Wöllstadt'],
}

# Gebiet entries to ignore during mapping; these are aggregated or irrelevant administrative areas
ignore_list = {'Ausgewählte Gebiete zusammengefasst', 'Sanierungsgebiet'}

def normalize(text: str) -> str:
    """
    Normalizes German strings to a simplified format for comparison.
    Converts umlauts and ß, lowercases, removes hyphens, trims whitespace.
    Used for matching Gebiet names more reliably across inconsistent data.
    """
    return (
        text.lower()
        .replace('ä', 'ae')
        .replace('ö', 'oe')
        .replace('ü', 'ue')
        .replace('ß', 'ss')
        .replace('-', ' ')
        .strip()
    )

def get_gemeinde_from_gebiet(gebiet: str) -> str:
    """
    Maps a given Gebiet name to its corresponding Gemeinde.
    - First checks if Gebiet is in the ignore list
    - Then checks for exact match in `mapping_gemeinde`
    - Then tries normalized fuzzy match against `mapping` entries
    - Logs warning if no match is found

    Returns the Gemeinde object (class instance), or an empty string if ignored.
    """
    if gebiet in ignore_list:
        return ''

    for gemeinde_var, gemeinde_names in mapping_gemeinde.items():
        if gebiet in gemeinde_names:
            return gemeinde_var

    norm_gebiet = normalize(gebiet)

    for gemeinde, gebieten in mapping.items():
        for teil_list in gebieten:
            if any(normalize(name) == norm_gebiet for name in teil_list):
                return gemeinde

    logging.warning(f"No Mapping found for: '{gebiet}'")
    return gebiet


def track_undetected_gebiete(input_gebiete: list[str]) -> dict:
    """
    Tracks which Gebiete were successfully matched to which Gemeinde.
    Used for post-validation and completeness checking.

    Returns a dictionary {Gemeinde: set of matched Gebiete}.
    """
    found_by_gemeinde = {}

    for gebiet in input_gebiete:
        gemeinde = get_gemeinde_from_gebiet(gebiet)
        if gemeinde is None:
            continue

        found_by_gemeinde.setdefault(gemeinde, set()).add(gebiet)

    return found_by_gemeinde

def log_missing_gebiete(found_by_gemeinde: dict):
    """
    Logs warnings for Gebiete that are expected (from `mapping`) but were not found in the dataset.
    Helps catch missing data or incorrect Gebiet names in source files.
    """
    for gemeinde, all_gebiete_nested in mapping.items():
        missing = []

        for gebiet_group in all_gebiete_nested:
            if not any(g in found_by_gemeinde.get(gemeinde, set()) for g in gebiet_group):
                missing.extend(gebiet_group)

        if missing:
            logging.warning(
                f"Gemeinde '{mapping_gemeinde[gemeinde][0]}' is missing {len(missing)} Gebiet(e): {sorted(missing)}"
            )

def get_gemeinde_by_schluessel(schluessel: str) -> str | None:
    """
    Maps an official Gemeindeschlüssel (key) to the canonical Gemeinde name.
    Logs a warning if the key is unknown.
    """
    if not schluessel:
        return None

    if schluessel in schluessel_map:
        return schluessel_map[schluessel]

    logging.warning(f"Gemeinde not found for schluessel: '{schluessel}'")
    return None

def normalize_gemeinde_name(raw_name: str) -> str:
    """
    Normalizes a raw Gemeinde name (from filename or data entry) to its canonical version.
    Uses `gemeinde_aliases` to map known variants.
    Logs a warning for any unmatched name.
    """
    norm = raw_name.strip().lower().replace("-", " ").replace("ß", "ss")

    for canonical, alternatives in gemeinde_aliases.items():
        for alt in alternatives:
            alt_norm = alt.strip().lower().replace("-", " ").replace("ß", "ss")
            if norm == alt_norm:
                return canonical

    logging.warning(
        f"Gemeinde mapping: unknown name '{raw_name}' — not mapped to canonical."
    )
    return raw_name