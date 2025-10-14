import pandas as pd

INPUT = "kartendarstellung.csv"
OUTPUT = "kartendarstellung_long.csv"

df = pd.read_csv(INPUT)

df = df.rename(columns={
    "Kommune": "gemeinde",
    "Schlüssel": "gemeinde_schluessel",
})

vars8 = [
    "Anzahl Bevölkerung je Kommune",
    "Fläche der Kommune",
    "Bevölkerungsdichte",
    "Unter 21 Jährige",
    "21 bis 65 Jährige",
    "Über 65 Jährige",
    "Jugendquotient",
    "Altenquotient",
]

df["iso"] = ""

long_df = df.melt(
    id_vars=["gemeinde", "gemeinde_schluessel", "iso", "jahr", "contour"],
    value_vars=vars8,
    var_name="variable",
    value_name="value",
)

long_df = long_df[[
    "gemeinde",
    "gemeinde_schluessel",
    "iso",
    "jahr",
    "contour",
    "variable",
    "value",
]]

long_df.to_csv(OUTPUT, index=False)
print(f"Saved: {OUTPUT}, rows: {len(long_df)}")
