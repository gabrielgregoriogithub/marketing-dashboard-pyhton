import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
import pandas as pd
import numpy as np

df = pd.read_csv(r"C:\Users\gabri\Downloads\deel\Deel Marketing Analytics Challenge - raw_data.csv",
                 sep=";", encoding="utf-8")
df.columns = df.columns.str.strip().str.replace('\xa0', ' ')

cost_raw = (df["$ Cost"].astype(str)
            .str.replace("$","",regex=False).str.replace("\xa0","",regex=False).str.strip())
mask_dot = cost_raw.str.contains(r"\.") & ~cost_raw.str.contains(",")
cost_raw.loc[mask_dot] = cost_raw.loc[mask_dot].str.replace(".", "", regex=False)
mask_comma = cost_raw.str.contains(",") & ~cost_raw.str.contains(r"\.")
cost_raw.loc[mask_comma] = cost_raw.loc[mask_comma].str.replace(",", ".", regex=False)
df["Cost"] = pd.to_numeric(cost_raw, errors="coerce").fillna(0)

for c in ["# Impressions","# Clicks","# Opportunities Won"]:
    df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

paid = df[df["Marketing Team"]=="paid"]
imp = (paid[paid["# Impressions"]>0]
       .groupby("Marketing Program")
       .agg(impressions=("# Impressions","sum"), clicks=("# Clicks","sum"),
            cost=("Cost","sum"), won=("# Opportunities Won","sum"))
       .reset_index())
imp["CTR"] = imp["clicks"]/imp["impressions"]*100
imp["CPM"] = imp["cost"]/imp["impressions"]*1000
imp["CPA"] = np.where(imp["won"]>0, imp["cost"]/imp["won"], np.nan)

print(f"\n{'Program':<28} {'Impressions':>14} {'Cost':>10} {'CPM':>10} {'CTR%':>8} {'CPA':>10}")
print("-"*82)
for _, r in imp.iterrows():
    cpa = f"${r['CPA']:.2f}" if not np.isnan(r['CPA']) else "n/a"
    print(f"{r['Marketing Program']:<28} {r['impressions']:>14,.0f} {r['cost']:>10,.2f}"
          f" {r['CPM']:>10.4f} {r['CTR']:>8.2f} {cpa:>10}")
