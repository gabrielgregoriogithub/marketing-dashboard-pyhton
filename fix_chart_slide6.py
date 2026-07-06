"""
Gera apenas o gráfico 'Monthly Spend vs Won' do Slide 6 com eixo Y corrigido.
Fix: remover a dupla divisão por 1000 — o formatter já cuida da conversão.
"""
import io, os
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker

# ── Cores ─────────────────────────────────────────────────────────────────────
MO = "#FF5C35"; MN = "#1E3A5F"; MM = "#6B7280"; MLG = "#F5F5F7"

# ── Dados ─────────────────────────────────────────────────────────────────────
def load():
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                        "Deel Marketing Analytics Challenge - raw_data.csv")
    df = pd.read_csv(path, sep=";", encoding="utf-8")
    df.columns = df.columns.str.strip().str.replace("\xa0", " ")
    df["Reporting Date"] = pd.to_datetime(df["Reporting Date"], dayfirst=True)

    cost_raw = (df["$ Cost"].astype(str)
                .str.replace("$", "", regex=False)
                .str.replace("\xa0", "", regex=False)
                .str.strip())
    mask_dot   = cost_raw.str.contains(r"\.") & ~cost_raw.str.contains(",")
    cost_raw.loc[mask_dot]   = cost_raw.loc[mask_dot].str.replace(".", "", regex=False)
    mask_comma = cost_raw.str.contains(",") & ~cost_raw.str.contains(r"\.")
    cost_raw.loc[mask_comma] = cost_raw.loc[mask_comma].str.replace(",", ".", regex=False)
    df["Cost"] = pd.to_numeric(cost_raw, errors="coerce").fillna(0)

    df["# Opportunities Won"] = pd.to_numeric(df["# Opportunities Won"], errors="coerce").fillna(0)
    df["YM"] = df["Reporting Date"].dt.to_period("M").astype(str)
    return df

# ── Gráfico ───────────────────────────────────────────────────────────────────
df = load()
m = (df.groupby("YM")
       .agg(won=("# Opportunities Won", "sum"),
            cost=("Cost", "sum"))
       .sort_values("YM")
       .reset_index())

x      = range(len(m))
labels = [v[2:7] for v in m["YM"]]

fig, ax2 = plt.subplots(figsize=(7, 3.8))
fig.patch.set_facecolor("white")

# BARRAS: custo em dollars brutos (sem dividir por 1000 aqui)
# O formatter abaixo cuida de exibir em $K
ax2.bar(x, m["cost"], color=MO, alpha=0.6, label="Spend ($K)")

# Eixo Y esquerdo — formatter divide por 1000 e exibe como $K
ax2.yaxis.set_major_formatter(
    mticker.FuncFormatter(lambda v, _: f"${v/1e3:.0f}K")
)

# Linha de Won (eixo direito)
ax2b = ax2.twinx()
ax2b.plot(x, m["won"], color=MN, lw=2, marker="o", ms=3, label="Won")
ax2b.spines[["top", "right"]].set_visible(False)
ax2b.tick_params(colors="#374151", labelsize=7)
ax2b.set_ylabel("Won Opps.", fontsize=8, color=MM)

# Eixo X
ax2.set_xticks(list(x)[::3])
ax2.set_xticklabels(labels[::3], rotation=45, ha="right", fontsize=7)

# Tema
ax2.set_facecolor(MLG)
ax2.spines[["top", "right"]].set_visible(False)
ax2.spines[["left", "bottom"]].set_color("#D1D5DB")
ax2.tick_params(colors="#374151", labelsize=7)
ax2.grid(axis="y", color="#E5E7EB", linewidth=0.6, linestyle="--")
ax2.set_title("Monthly Spend vs Won", fontsize=11, fontweight="bold",
              color="#141423", pad=8)
ax2.set_ylabel("Monthly Spend", fontsize=8, color=MM)

# Legenda combinada
h1, l1 = ax2.get_legend_handles_labels()
h2, l2 = ax2b.get_legend_handles_labels()
ax2.legend(h1 + h2, l1 + l2, fontsize=8, loc="upper left")

fig.tight_layout(pad=1.5)

out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "chart_slide6_fixed.png")
fig.savefig(out, dpi=180, bbox_inches="tight")
print(f"Salvo em: {out}")
