import pandas as pd

df1 = pd.read_excel("sql-lab-test-1.xlsx")
df2 = pd.read_excel("sql-lab-test-2.xlsx")

merged = pd.merge(df1, df2, on=["Gemeinde", "Jahr"])
merged.to_excel("merged.xlsx", sheet_name="Merged", index=False)