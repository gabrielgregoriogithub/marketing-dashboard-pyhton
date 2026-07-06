import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

import pandas as pd
import numpy as np

df = pd.read_csv(
    r"C:\Users\gabri\Downloads\deel\Deel Marketing Analytics Challenge - raw_data.csv",
    sep=";", encoding="utf-8"
)
df.columns = df.columns.str.strip().str.replace('\xa0', ' ')

# Cost parsing identical to app.py
cost_raw = (df["$ Cost"].astype(str)
            .str.replace("$", "", regex=False)
            .str.replace("\xa0", "", regex=False)
            .str.strip())
mask_dot = cost_raw.str.contains(r"\.") & ~cost_raw.str.contains(",")
cost_raw.loc[mask_dot] = cost_raw.loc[mask_dot].str.replace(".", "", regex=False)
mask_comma = cost_raw.str.contains(",") & ~cost_raw.str.contains(r"\.")
cost_raw.loc[mask_comma] = cost_raw.loc[mask_comma].str.replace(",", ".", regex=False)
df["Cost"] = pd.to_numeric(cost_raw, errors="coerce").fillna(0)

# Cost_sim parsing (correct real-world values)
cost_sim = (df["$ Cost"].astype(str)
            .str.replace("$", "", regex=False)
            .str.replace("\xa0", "", regex=False)
            .str.strip())
mask_eu = cost_sim.str.count(r"\.") > 1
cost_sim.loc[mask_eu] = cost_sim.loc[mask_eu].str.replace(".", "", regex=False)
cost_sim = cost_sim.str.replace(",", "", regex=False)
df["Cost_sim"] = pd.to_numeric(cost_sim, errors="coerce").fillna(0)

paid = df[df["Marketing Team"] == "paid"]

summary = (paid.groupby("Marketing Program")
           .agg(
               cost=("Cost", "sum"),
               cost_sim=("Cost_sim", "sum"),
               won=("# Opportunities Won", "sum"),
               rows=("$ Cost", "count")
           ).reset_index())

summary["CPA_cost"]     = np.where(summary["won"] > 0, summary["cost"]     / summary["won"], np.nan)
summary["CPA_cost_sim"] = np.where(summary["won"] > 0, summary["cost_sim"] / summary["won"], np.nan)

summary = summary.sort_values("CPA_cost")

print(f"\n{'Program':<25} {'Cost':>12} {'Cost_sim':>14} {'Won':>8} {'CPA(Cost)':>12} {'CPA(Sim)':>12}")
print("-" * 85)
for _, r in summary.iterrows():
    cpa1 = f"${r['CPA_cost']:,.2f}"     if not np.isnan(r['CPA_cost'])     else "n/a"
    cpa2 = f"${r['CPA_cost_sim']:,.0f}" if not np.isnan(r['CPA_cost_sim']) else "n/a"
    print(f"{r['Marketing Program']:<25} {r['cost']:>12,.2f} {r['cost_sim']:>14,.0f} "
          f"{r['won']:>8,.2f} {cpa1:>12} {cpa2:>12}")

print(f"\nTotal Cost:     {summary['cost'].sum():,.2f}")
print(f"Total Cost_sim: {summary['cost_sim'].sum():,.0f}")

# Show raw $ Cost samples for awareness and sponsorship
for prog in ["paid_awareness", "paid_sponsorship"]:
    sub = paid[paid["Marketing Program"] == prog][["$ Cost","Cost","Cost_sim","# Opportunities Won"]]
    print(f"\n--- {prog} (first 5 rows) ---")
    print(sub.head(5).to_string())
    print(f"  Won total: {sub['# Opportunities Won'].sum():.2f}")
    print(f"  Cost total: {sub['Cost'].sum():.4f}")
    print(f"  Cost_sim total: {sub['Cost_sim'].sum():,.0f}")
