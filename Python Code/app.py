import os
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st
from pathlib import Path

# ── Config ────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Deel Marketing Analytics",
    layout="wide",
    initial_sidebar_state="expanded",
)

C = {
    "blue": "#4a9eed", "amber": "#f59e0b", "green": "#22c55e",
    "red": "#ef4444", "purple": "#8b5cf6", "pink": "#ec4899",
    "cyan": "#06b6d4", "lime": "#84cc16", "navy": "#1e3a5f", "bg": "#f8f9fa",
}

MONTHS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]

# ── Data ──────────────────────────────────────────────────────────────────────



@st.cache_data
def load_data() -> pd.DataFrame:
    path = Path(__file__).parent / "Deel Marketing Analytics Challenge - raw_data.csv"
    df = pd.read_csv(path, sep=";", encoding="utf-8")
    df.columns = df.columns.str.strip().str.replace("\xa0", " ")
    df["Reporting Date"] = pd.to_datetime(df["Reporting Date"], dayfirst=True)


    # COST
    cost = (
    df["$ Cost"]
    .astype(str)
    .str.replace("$", "", regex=False)
    .str.replace("\xa0", "", regex=False)
    .str.strip()
    )


    mask_dot = cost.str.contains(r"\.") & ~cost.str.contains(",")

    cost.loc[mask_dot] = cost.loc[mask_dot].str.replace(".", "", regex=False)

    # caso decimal: 160,756
    mask_comma = cost.str.contains(",") & ~cost.str.contains(r"\.")

    cost.loc[mask_comma] = cost.loc[mask_comma].str.replace(",", ".", regex=False)

    df["Cost"] = pd.to_numeric(cost, errors="coerce")

    cost_sim = (
        df["$ Cost"]
        .astype(str)
        .str.replace("$", "", regex=False)
        .str.replace("\xa0", "", regex=False)
        .str.strip()
    )
    mask_eu = cost_sim.str.count(r"\.") > 1
    cost_sim.loc[mask_eu] = cost_sim.loc[mask_eu].str.replace(".", "", regex=False)
    cost_sim = cost_sim.str.replace(",", "", regex=False)
    df["Cost_sim"] = pd.to_numeric(cost_sim, errors="coerce").fillna(0)

    for c in [
        "# Prospects", "# Marketing Qualified Prospects", "# Demo Prospects",
        "# Sales Qualified Opportunities", "# Opportunities Won",
        "# Impressions", "# Clicks",
    ]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    df["YM"]       = df["Reporting Date"].dt.to_period("M").astype(str)
    df["Year"]     = df["Reporting Date"].dt.year
    df["MonthNum"] = df["Reporting Date"].dt.month
    df["MonthAbb"] = df["Reporting Date"].dt.strftime("%b")
    df["Cycle"] = df["Reporting Date"].apply(
        lambda d: "Jul 23-Jun 24"
        if (d.year == 2023 and d.month >= 7) or (d.year == 2024 and d.month <= 6)
        else (
            "Jul 24-Jun 25"
            if (d.year == 2024 and d.month >= 7) or (d.year == 2025 and d.month <= 6)
            else "Outros"
        )
    )
    return df

# ── UI helpers ────────────────────────────────────────────────────────────────
def kpi(col, label, value, sub=None, color=C["blue"]):
    with col:
        sub_html = (
            f"<div style='font-size:12px;color:#6b7280;margin-top:3px'>{sub}</div>"
            if sub else ""
        )
        st.markdown(
            f"""<div style="border-left:4px solid {color};padding:14px 16px;
            background:#fff;border-radius:8px;box-shadow:0 1px 4px rgba(0,0,0,.07);
            min-height:88px">
            <div style="font-size:11px;color:#9ca3af;text-transform:uppercase;
            letter-spacing:.6px">{label}</div>
            <div style="font-size:26px;font-weight:700;color:#111827;
            line-height:1.2">{value}</div>
            {sub_html}</div>""",
            unsafe_allow_html=True,
        )

def section(title: str):
    st.markdown(
        f"<h4 style='margin:20px 0 6px;color:{C['navy']};font-size:15px;"
        f"border-bottom:2px solid {C['blue']};padding-bottom:5px'>{title}</h4>",
        unsafe_allow_html=True,
    )

def theme(fig, h=360, legend=True):
    fig.update_layout(
        height=h,
        plot_bgcolor=C["bg"],
        paper_bgcolor="white",
        font=dict(family="Inter, Arial, sans-serif", size=12, color="#374151"),
        margin=dict(l=30, r=20, t=38, b=28),
        showlegend=legend,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    )
    fig.update_xaxes(gridcolor="#e5e7eb", zerolinecolor="#d1d5db", linecolor="#e5e7eb")
    fig.update_yaxes(gridcolor="#e5e7eb", zerolinecolor="#d1d5db", linecolor="#e5e7eb")
    return fig

# ── Page 1: Visao Executiva ───────────────────────────────────────────────────
def p1_executive(df):
    tc  = df["Cost"].sum()
    tw  = df["# Opportunities Won"].sum()
    tp  = df["# Prospects"].sum()
    cpa = tc / tw if tw else 0
    e2e = tw / tp * 100 if tp else 0

    c1, c2, c3, c4 = st.columns(4)
    kpi(c1, "Total Invested",    f"${tc/1e3:.1f}K",  color=C["amber"])
    kpi(c2, "Won Opps.",         f"{tw:,.0f}",       color=C["green"])
    kpi(c3, "Avg. Global CPA",   f"${cpa:,.0f}",     color=C["blue"])
    kpi(c4, "E2E Rate",          f"{e2e:.2f}%",      color=C["purple"])

    st.markdown("<br>", unsafe_allow_html=True)

    l, r = st.columns(2)

    # ── Funil ────────────────────────────────────────────────
    with l:
        section("Global Funnel — Volumes & Drop-off")

        stages = ["Prospects", "MQPs", "Demos", "SQOs", "Won"]
        vals = [
            tp,
            df["# Marketing Qualified Prospects"].sum(),
            df["# Demo Prospects"].sum(),
            df["# Sales Qualified Opportunities"].sum(),
            tw,
        ]

        # % em relação ao topo
        pct_top = [
            (v / vals[0] * 100) if vals[0] > 0 else 0
            for v in vals
        ]

        # % de conversão entre etapas
        pct_stage = [100]
        for i in range(1, len(vals)):
            pct_stage.append(
                (vals[i] / vals[i - 1] * 100)
                if vals[i - 1] > 0 else 0
            )

        fig = go.Figure(
            go.Funnel(
              y=stages,
              x=vals,
              text=[
                 f"{v:,.0f}<br>({p:.1f}%)'"
                 for v, p in zip(vals, pct_top)
             ],
             textinfo="text",
             textposition="auto",   # <- importante
             marker=dict(
              color=[
                    C["navy"],
                  "#2563eb",
                   C["blue"],
                  "#93c5fd",
                  C["green"],
            ]
        ),
            connector=dict(
              line=dict(
                    color="#9ca3af",
                    dash="dot",
                    width=1,
                )
           ),
        )
    )

        max_val = max(vals)

        # Legenda das conversões
        fig.add_annotation(
            x=-max_val * 0.12,
            y=-0.05,
            xref="x",
            yref="paper",
            text="<b>Stage Conversion</b>",
            showarrow=False,
            font=dict(
                size=11,
                color="#dc2626",
            ),
            align="left",
        )

        # Conversões entre etapas
        for i in range(1, len(stages)):
            fig.add_annotation(
                x=-max_val * 0.12,
                y=stages[i],
                text=f"<b>{pct_stage[i]:.1f}%</b>",
                showarrow=False,
                font=dict(
                    size=12,
                    color="#dc2626",
                ),
            )

        fig.update_layout(
            margin=dict(
                l=120,
                r=20,
                t=35,
                b=50,
            )
        )

        theme(fig, h=340, legend=False)
        st.plotly_chart(fig, use_container_width=True)

  # ── Treemap ──────────────────────────────────────────────
    with r:
        section("Investment by Program (size = spend | color = CPA)")

        pp = (
            df[df["Cost"] > 0]
            .groupby("Marketing Program")
            .agg(
                cost=("Cost", "sum"),
                won=("# Opportunities Won", "sum"),
            )
            .reset_index()
        )

        pp["CPA"] = np.where(
            pp["won"] > 0,
            pp["cost"] / pp["won"],
            pp["cost"],
        )

        fig2 = px.treemap(
            pp,
            path=["Marketing Program"],
            values="cost",
            color="CPA",
            color_continuous_scale=[
                "#22c55e",
                "#f59e0b",
                "#ef4444",
            ],
            custom_data=["won", "CPA"],
        )

        fig2.update_traces(
            texttemplate="<b>%{label}</b><br>$%{value:,.0f}",
            hovertemplate=(
                "<b>%{label}</b><br>"
                "Spend: $%{value:,.0f}<br>"
                "Won: %{customdata[0]:,.0f}<br>"
                "CPA: $%{customdata[1]:,.0f}"
                "<extra></extra>"
            ),
        )

        theme(fig2, h=340, legend=False)
        st.plotly_chart(fig2, use_container_width=True)

    section("Monthly Trend — Won Opps. vs Investment")

    monthly = (
        df.groupby("YM")
        .agg(
            won=("# Opportunities Won", "sum"),
            cost=("Cost", "sum"),
        )
        .reset_index()
        .sort_values("YM")
    )

    fig3 = make_subplots(
        specs=[[{"secondary_y": True}]]
    )

    fig3.add_trace(
        go.Scatter(
            x=monthly["YM"],
            y=monthly["won"],
            name="Won Opps.",
            fill="tozeroy",
            line=dict(
                color=C["green"],
                width=2,
            ),
            fillcolor="rgba(34,197,94,.12)",
        ),
        secondary_y=False,
    )

    fig3.add_trace(
        go.Scatter(
            x=monthly["YM"],
            y=monthly["cost"],
            name="Investment ($)",
            line=dict(
                color=C["amber"],
                width=2,
                dash="dot",
            ),
        ),
        secondary_y=True,
    )

    fig3.update_yaxes(
        title_text="Won Opps.",
        secondary_y=False,
        gridcolor="#e5e7eb",
    )

    fig3.update_yaxes(
        title_text="Investment ($)",
        secondary_y=True,
        gridcolor="rgba(0,0,0,0)",
    )

    theme(fig3, h=290)
    st.plotly_chart(fig3, use_container_width=True)

# ── Page 2: Funil Completo ────────────────────────────────────────────────────
def p2_funnel(df):
    by_t = (
        df.groupby("Marketing Team")
        .agg(
            prospects=("# Prospects", "sum"),
            mqp=("# Marketing Qualified Prospects", "sum"),
            demo=("# Demo Prospects", "sum"),
            sqo=("# Sales Qualified Opportunities", "sum"),
            won=("# Opportunities Won", "sum"),
            cost=("Cost", "sum"),
        )
        .reset_index()
    )
    safe = lambda a, b: np.where(b > 0, a / b * 100, np.nan)
    by_t["MQP%"]  = safe(by_t["mqp"],  by_t["prospects"]).round(1)
    by_t["Demo%"] = safe(by_t["demo"], by_t["mqp"]).round(1)
    by_t["SQO%"]  = safe(by_t["sqo"],  by_t["demo"]).round(1)
    by_t["Win%"]  = safe(by_t["won"],  by_t["sqo"]).round(1)
    by_t["E2E%"]  = safe(by_t["won"],  by_t["prospects"]).round(2)
    by_t["CPA"]   = np.where(by_t["cost"] > 0, by_t["cost"] / by_t["won"].replace(0, np.nan), 0)

    l, r = st.columns(2)

    with l:
        section("Conversion Rates by Stage — All Teams")
        stage_colors = {
            "MQP%": C["blue"], "Demo%": C["cyan"],
            "SQO%": C["amber"], "Win%": C["green"], "E2E%": C["purple"],
        }
        fig = go.Figure()
        for s, col in stage_colors.items():
            fig.add_trace(go.Bar(
                name=s, x=by_t["Marketing Team"], y=by_t[s], marker_color=col,
            ))
        fig.update_layout(barmode="group")
        theme(fig, h=360)
        st.plotly_chart(fig, use_container_width=True)

    with r:
        section("Efficiency Table by Team")
        disp = by_t[["Marketing Team", "MQP%", "Demo%", "SQO%", "Win%", "E2E%", "CPA"]].copy()
        disp = disp.sort_values("E2E%", ascending=False).reset_index(drop=True)
        rows_html = ""
        for _, row in disp.iterrows():
            cpa_str = "—" if row["CPA"] == 0 else f"${row['CPA']:,.0f}"
            e2e_color = C["green"] if row["E2E%"] >= 2 else (C["amber"] if row["E2E%"] >= 0.5 else C["red"])
            rows_html += (
                f"<tr>"
                f"<td style='padding:7px 10px;border-bottom:1px solid #f3f4f6;"
                f"font-weight:500'>{row['Marketing Team']}</td>"
                f"<td style='text-align:center;padding:7px 8px;border-bottom:1px solid #f3f4f6'>"
                f"{row['MQP%']:.1f}%</td>"
                f"<td style='text-align:center;padding:7px 8px;border-bottom:1px solid #f3f4f6'>"
                f"{row['Demo%']:.1f}%</td>"
                f"<td style='text-align:center;padding:7px 8px;border-bottom:1px solid #f3f4f6'>"
                f"{row['SQO%']:.1f}%</td>"
                f"<td style='text-align:center;padding:7px 8px;border-bottom:1px solid #f3f4f6'>"
                f"{row['Win%']:.1f}%</td>"
                f"<td style='text-align:center;padding:7px 8px;border-bottom:1px solid #f3f4f6;"
                f"color:{e2e_color};font-weight:700'>{row['E2E%']:.2f}%</td>"
                f"<td style='text-align:center;padding:7px 8px;border-bottom:1px solid #f3f4f6'>"
                f"{cpa_str}</td>"
                f"</tr>"
            )
        st.markdown(
            f"""<div style='overflow-x:auto;margin-top:8px'>
            <table style='width:100%;border-collapse:collapse;font-size:13px'>
            <thead><tr style='background:{C["navy"]};color:white'>
            <th style='padding:8px 10px;text-align:left'>Team</th>
            <th style='padding:8px 8px'>MQP%</th><th style='padding:8px 8px'>Demo%</th>
            <th style='padding:8px 8px'>SQO%</th><th style='padding:8px 8px'>Win%</th>
            <th style='padding:8px 8px'>E2E%</th><th style='padding:8px 8px'>CPA</th>
            </tr></thead><tbody>{rows_html}</tbody></table></div>""",
            unsafe_allow_html=True,
        )

    st.markdown("<br>", unsafe_allow_html=True)

    with st.container():
        section("E2E% Heatmap by Team × Month (last 18 months)")
        hdf = (
            df.groupby(["YM", "Marketing Team"])
            .agg(p=("# Prospects", "sum"), w=("# Opportunities Won", "sum"))
            .reset_index()
        )
        hdf["E2E"] = np.where(hdf["p"] > 0, hdf["w"] / hdf["p"] * 100, np.nan)
        pivot = hdf.pivot_table(index="Marketing Team", columns="YM", values="E2E", aggfunc="mean")
        pivot = pivot.iloc[:, -18:]
        fig = go.Figure(go.Heatmap(
            z=pivot.values,
            x=pivot.columns.tolist(),
            y=pivot.index.tolist(),
            colorscale=[[0, "#fef3c7"], [0.5, C["blue"]], [1, C["navy"]]],
            text=np.where(np.isnan(pivot.values), "",
                          np.round(pivot.values, 1).astype(str)),
            texttemplate="%{text}",
            textfont=dict(size=9),
            hoverongaps=False,
            colorbar=dict(title="E2E%"),
        ))
        theme(fig, h=280, legend=False)
        st.plotly_chart(fig, use_container_width=True)

    section("Scatter: Prospect Volume vs E2E Rate by Program")
    sp = (
        df.groupby("Marketing Program")
        .agg(pros=("# Prospects", "sum"), won=("# Opportunities Won", "sum"),
             cost=("Cost", "sum"))
        .reset_index()
    )
    sp["E2E"] = np.where(sp["pros"] > 0, sp["won"] / sp["pros"] * 100, 0)
    sp["sz"]  = np.where(sp["cost"] > 0, sp["cost"] / 1e6, 0.3)
    sp = sp[sp["pros"] > 100]
    fig = px.scatter(
        sp, x="pros", y="E2E", size="sz", color="Marketing Program",
        text="Marketing Program",
        labels={"pros": "Prospects", "E2E": "E2E Rate (%)", "sz": "Spend ($M)"},
        size_max=80,
        color_discrete_sequence=px.colors.qualitative.Bold,
    )
    fig.update_traces(textposition="top center", textfont_size=10)
    theme(fig, h=450)
    st.plotly_chart(fig, use_container_width=True)


# ── Page 3: Canais Pagos ──────────────────────────────────────────────────────
def p3_paid(df):
    paid = df[df["Marketing Team"] == "paid"]
    top5 = (
        paid.groupby("Marketing Program")["Cost"]
        .sum()
        .sort_values(ascending=False)
        .head(5)
        .index.tolist()
    )

    section("Monthly CPA — Top Paid Programs")
    cpa_t = (
        paid.groupby(["YM", "Marketing Program"])
        .agg(cost=("Cost", "sum"), won=("# Opportunities Won", "sum"))
        .reset_index()
    )
    cpa_t = cpa_t[cpa_t["Marketing Program"].isin(top5)]
    cpa_t["CPA"] = np.where(cpa_t["won"] > 0, cpa_t["cost"] / cpa_t["won"], np.nan)
    cpa_t["CPA_K"] = cpa_t["CPA"]
    fig = px.line(
        cpa_t.dropna(subset=["CPA_K"]).sort_values("YM"),
        x="YM", y="CPA_K", color="Marketing Program",
        labels={"YM": "Month", "CPA_K": "CPA ($)", "Marketing Program": "Program"},
        color_discrete_sequence=[C["blue"], C["green"], C["amber"], C["red"], C["purple"]],
    )
    fig.update_traces(line_width=2)
    fig.add_hline(
        y=1056, line_dash="dot", line_color="#9ca3af",
        annotation_text="Benchmark $1056", annotation_position="bottom right",
        annotation_font_size=11,
    )
    theme(fig, h=310)
    fig.update_yaxes(tickprefix="$", ticksuffix="")
    st.plotly_chart(fig, use_container_width=True)

    l, r = st.columns(2)

    with l:
        section("CPM vs CTR Matrix  (bubble size = CPA)")
        imp = (
            paid[paid["# Impressions"] > 0]
            .groupby("Marketing Program")
            .agg(
                impressions=("# Impressions", "sum"),
                clicks=("# Clicks", "sum"),
                cost=("Cost", "sum"),
                won=("# Opportunities Won", "sum"),
            )
            .reset_index()
        )
        imp["CTR"] = imp["clicks"] / imp["impressions"] * 100
        imp["CPM"] = imp["cost"]   / imp["impressions"] * 1000
        imp["CPA"] = np.where(imp["won"] > 0, imp["cost"] / imp["won"], np.nan)
        fig = px.scatter(
            imp.dropna(subset=["CPA"]),
            x="CPM", y="CTR", size="CPA", text="Marketing Program",
            color="CPA",
            color_continuous_scale=["#22c55e", "#f59e0b", "#ef4444"],
            labels={"CPM": "CPM ($)", "CTR": "CTR (%)", "CPA": "CPA ($)"},
            size_max=60,
        )
        fig.update_traces(textposition="top center", textfont_size=9)
        fig.add_hline(y=imp["CTR"].median(), line_dash="dash", line_color="#d1d5db")
        fig.add_vline(x=imp["CPM"].median(), line_dash="dash", line_color="#d1d5db")
        theme(fig, h=360, legend=False)
        st.plotly_chart(fig, use_container_width=True)

    with r:
        section("CPA by Program  (green = below $1K)")
        cb = (
            paid.groupby("Marketing Program")
            .agg(cost=("Cost", "sum"), won=("# Opportunities Won", "sum"))
            .reset_index()
        )
        cb["CPA"] = np.where(cb["won"] > 0, cb["cost"] / cb["won"], np.nan)
        cb = cb.dropna(subset=["CPA"]).sort_values("CPA")
        fig = go.Figure(
    go.Bar(
        x=cb["CPA"],
        y=cb["Marketing Program"],
        orientation="h",
        marker_color=[
            C["green"] if v <= 1000 else C["red"]
            for v in cb["CPA"]
        ],
        text=[f"${v:,.0f}" for v in cb["CPA"]],
        textposition="auto",   # <- em vez de outside
    )
)
        fig.add_vline(
            x=1000, line_dash="dot", line_color="#9ca3af",
            annotation_text="$1.000", annotation_position="top",
        )
        theme(fig, h=360, legend=False)
        st.plotly_chart(fig, use_container_width=True)

    section("Spend vs Won by Fiscal Cycle — Top 5 Programs")
    cyc = (
        paid[paid["Cycle"] != "Outros"]
        .groupby(["Cycle", "Marketing Program"])
        .agg(cost=("Cost", "sum"), won=("# Opportunities Won", "sum"))
        .reset_index()
    )
    cyc = cyc[cyc["Marketing Program"].isin(top5)]
    l2, r2 = st.columns(2)
    cmap = {"Jul 23-Jun 24": C["blue"], "Jul 24-Jun 25": C["navy"]}
    with l2:
        fig = px.bar(cyc, x="Marketing Program", y="cost", color="Cycle",
                     barmode="group", color_discrete_map=cmap,
                     labels={"cost": "Spend ($)", "Marketing Program": "Program"},
                     title="Spend YoY")
        theme(fig, h=280)
        st.plotly_chart(fig, use_container_width=True)
    with r2:
        fig = px.bar(cyc, x="Marketing Program", y="won", color="Cycle",
                     barmode="group", color_discrete_map=cmap,
                     labels={"won": "Won Opps.", "Marketing Program": "Program"},
                     title="Won YoY")
        theme(fig, h=280)
        st.plotly_chart(fig, use_container_width=True)


# ── Page 4: Canais Zero-Custo ─────────────────────────────────────────────────
def p4_organic(df):
    org = df[df["Cost"] == 0]
    paid_cpa = df[df["Cost"] > 0]["Cost"].sum() / max(
        df[df["Cost"] > 0]["# Opportunities Won"].sum(), 1
    )

    lc  = org[org["Marketing Team"] == "lifecycle"]["# Opportunities Won"].sum()
    cnt = org[org["Marketing Team"] == "content"]["# Opportunities Won"].sum()
    web = org[org["Marketing Team"] == "website"]["# Opportunities Won"].sum()
    total_org = lc + cnt + web + org[org["Marketing Team"] == "social"]["# Opportunities Won"].sum()
    implied = total_org * paid_cpa

    c1, c2, c3, c4 = st.columns(4)
    kpi(c1, "Lifecycle Won",   f"{lc:,.0f}",           sub="+132% YoY",       color=C["purple"])
    kpi(c2, "Content Won",    f"{cnt:,.0f}",           sub="Zero cost",            color=C["amber"])
    kpi(c3, "Website Won",    f"{web:,.0f}",           sub="Zero cost",            color=C["green"])
    kpi(c4, "Implied Value",  f"${implied/1e6:.0f}M",  sub=f"@ paid CPA ${paid_cpa:,.0f}", color=C["cyan"])

    st.markdown("<br>", unsafe_allow_html=True)
    l, r = st.columns(2)

    with l:
        section("Won by Organic Channel — YoY Comparison")
        yoy = (
            org[org["Cycle"] != "Outros"]
            .groupby(["Cycle", "Marketing Team"])
            .agg(won=("# Opportunities Won", "sum"))
            .reset_index()
        )
        fig = px.bar(
            yoy, x="Marketing Team", y="won", color="Cycle",
            barmode="group",
            color_discrete_map={"Jul 23-Jun 24": C["blue"], "Jul 24-Jun 25": C["navy"]},
            labels={"won": "Won Opps.", "Marketing Team": "Team"},
        )
        theme(fig, h=320)
        st.plotly_chart(fig, use_container_width=True)

    with r:
        section("Monthly Trend — Lifecycle vs Website")
        lw = (
            df[df["Marketing Team"].isin(["lifecycle", "website"])]
            .groupby(["YM", "Marketing Team"])
            .agg(won=("# Opportunities Won", "sum"))
            .reset_index()
            .sort_values("YM")
        )
        fig = px.line(
            lw, x="YM", y="won", color="Marketing Team",
            color_discrete_map={"lifecycle": C["purple"], "website": C["green"]},
            markers=True,
            labels={"YM": "Month", "won": "Won Opps.", "Marketing Team": "Team"},
        )
        fig.update_traces(line_width=2, marker_size=5)
        theme(fig, h=320)
        st.plotly_chart(fig, use_container_width=True)

    section("Efficiency by Organic Program")
    op = (
        org.groupby(["Marketing Team", "Marketing Program"])
        .agg(
            pros=("# Prospects", "sum"),
            won=("# Opportunities Won", "sum"),
        )
        .reset_index()
    )
    op["E2E%"]              = np.where(op["pros"] > 0, op["won"] / op["pros"] * 100, 0).round(2)
    op["Implied Value ($)"] = (op["won"] * paid_cpa).round(0).astype(int)
    op = op[op["won"] > 0].sort_values("E2E%", ascending=False).reset_index(drop=True)
    op["Implied Value ($)"] = op["Implied Value ($)"].apply(lambda x: f"${x:,}")
    op.columns = ["Team", "Program", "Prospects", "Won", "E2E%", "Implied Value"]
    st.dataframe(op, use_container_width=True, hide_index=True)


# ── Page 5: Tendencias & Sazonalidade ────────────────────────────────────────
def p5_trends(df):
    month_label = {i: MONTHS[i - 1] for i in range(1, 13)}
    fiscal_order = list(range(7, 13)) + list(range(1, 7))
    fiscal_pos   = {m: i + 1 for i, m in enumerate(fiscal_order)}

    section("YoY Trend — Won & Investment by Fiscal Cycle")
    yoy = (
        df[df["Cycle"] != "Outros"]
        .groupby(["Cycle", "MonthNum"])
        .agg(won=("# Opportunities Won", "sum"), cost=("Cost", "sum"))
        .reset_index()
    )
    yoy["FP"]    = yoy["MonthNum"].map(fiscal_pos)
    yoy["Label"] = yoy["MonthNum"].map(month_label)
    yoy = yoy.sort_values(["Cycle", "FP"])
    cmap = {"Jul 23-Jun 24": C["blue"], "Jul 24-Jun 25": C["navy"]}

    l, r = st.columns(2)
    with l:
        fig = px.line(
            yoy, x="FP", y="won", color="Cycle", markers=True,
            color_discrete_map=cmap,
            labels={"FP": "Fiscal Month", "won": "Won Opps."},
            title="Won by Month (Jul=1 ... Jun=12)",
        )
        fig.update_xaxes(
            tickmode="array",
            tickvals=list(range(1, 13)),
            ticktext=[month_label[m] for m in fiscal_order],
        )
        fig.update_traces(line_width=2, marker_size=6)
        theme(fig, h=310)
        st.plotly_chart(fig, use_container_width=True)
    with r:
        fig = px.line(
            yoy, x="FP", y="cost", color="Cycle", markers=True,
            color_discrete_map={"Jul 23-Jun 24": C["amber"], "Jul 24-Jun 25": C["red"]},
            labels={"FP": "Fiscal Month", "cost": "Investment ($)"},
            title="Investment by Month",
        )
        fig.update_xaxes(
            tickmode="array",
            tickvals=list(range(1, 13)),
            ticktext=[month_label[m] for m in fiscal_order],
        )
        fig.update_traces(line_width=2, marker_size=6)
        theme(fig, h=310)
        st.plotly_chart(fig, use_container_width=True)

    l2, r2 = st.columns(2)

    with l2:
        section("Seasonality — Won Heatmap (Year × Month)")
        sea = (
            df.groupby(["Year", "MonthNum"])
            .agg(won=("# Opportunities Won", "sum"))
            .reset_index()
        )
        sea["Mon"] = sea["MonthNum"].map(month_label)
        piv = sea.pivot_table(index="Year", columns="Mon", values="won", aggfunc="sum")
        ordered_cols = [MONTHS[i] for i in range(12) if MONTHS[i] in piv.columns]
        piv = piv.reindex(columns=ordered_cols)
        fig = go.Figure(go.Heatmap(
            z=piv.values,
            x=piv.columns.tolist(),
            y=piv.index.astype(str).tolist(),
            colorscale=[[0, "#f0f9ff"], [1, C["navy"]]],
            text=np.round(piv.values, 0),
            texttemplate="%{text:,.0f}",
            textfont=dict(size=11),
            colorbar=dict(title="Won"),
        ))
        theme(fig, h=230, legend=False)
        st.plotly_chart(fig, use_container_width=True)

    with r2:
        section("Monthly CPA — Efficiency Trend")
        cm = (
            df[df["Cost"] > 0]
            .groupby("YM")
            .agg(cost=("Cost", "sum"), won=("# Opportunities Won", "sum"))
            .reset_index()
            .sort_values("YM")
        )
        cm["CPA"] = cm["cost"] / cm["won"].replace(0, np.nan)
        cm["CPA_K"] = cm["CPA"]
        fig = px.line(
            cm.dropna(subset=["CPA_K"]), x="YM", y="CPA_K",
            labels={"YM": "Month", "CPA_K": "CPA ($)"},
            color_discrete_sequence=[C["blue"]],
        )
        fig.update_traces(line_width=2)
        fig.add_hrect(y0=0,  y1=35,   fillcolor=C["green"], opacity=0.07, line_width=0,
                      annotation_text="Efficient zone", annotation_position="bottom right",
                      annotation_font_size=10)
        fig.add_hrect(y0=35, y1=99, fillcolor=C["red"],   opacity=0.04, line_width=0,
                      annotation_text="Alert zone",     annotation_position="top right",
                      annotation_font_size=10)
        theme(fig, h=230, legend=False)
        fig.update_yaxes(tickprefix="$", tickformat=",.0f")
        st.plotly_chart(fig, use_container_width=True)

    section("YoY Growth by Team (Won  Jul 24-Jun 25 vs Jul 23-Jun 24)")
    gt = (
        df[df["Cycle"] != "Outros"]
        .groupby(["Cycle", "Marketing Team"])
        .agg(won=("# Opportunities Won", "sum"))
        .reset_index()
    )
    piv_g = gt.pivot_table(index="Marketing Team", columns="Cycle", values="won", aggfunc="sum")
    if "Jul 23-Jun 24" in piv_g.columns and "Jul 24-Jun 25" in piv_g.columns:
        piv_g["growth"] = (
            (piv_g["Jul 24-Jun 25"] - piv_g["Jul 23-Jun 24"]) / piv_g["Jul 23-Jun 24"] * 100
        ).round(1)
        piv_g = piv_g.sort_values("growth", ascending=False).reset_index()
        fig = go.Figure(go.Bar(
            x=piv_g["Marketing Team"],
            y=piv_g["growth"],
            marker_color=[C["green"] if v >= 0 else C["red"] for v in piv_g["growth"]],
            text=[f"{v:+.1f}%" for v in piv_g["growth"]],
            textposition="outside",
        ))
        fig.add_hline(y=0, line_color="#9ca3af")
        theme(fig, h=260, legend=False)
        st.plotly_chart(fig, use_container_width=True)


# ── Page 6: Scenario Simulator ────────────────────────────────────────────────
def p6_simulator(df):
    paid = df[df["Marketing Team"] == "paid"]

    hist = (
        paid.groupby("Marketing Program")
        .agg(
            cost_sim=("Cost_sim", "sum"),
            won=("# Opportunities Won", "sum")
        )
        .reset_index()
    )

    hist["CPA"] = np.where(
        hist["won"] > 0,
        hist["cost_sim"] / hist["won"],
        np.nan
    )

    hist = (
        hist.dropna(subset=["CPA"])
        .sort_values("cost_sim", ascending=False)
        .reset_index(drop=True)
    )

    total_budget = hist["cost_sim"].sum()
    cur_won = hist["won"].sum()

    st.markdown(
        f"""<div style="background:{C['navy']};color:white;padding:14px 20px;
        border-radius:8px;margin-bottom:18px">
        <b>Budget Reallocation Simulator</b> — Adjust the sliders to redistribute the
        <b>$421.5M</b> budget across paid programs and see the projected
        impact on won opportunities (based on each program's historical CPA).
        </div>""",
        unsafe_allow_html=True,
    )

    l, r = st.columns([1.3, 1])
    new_budgets = {}

    with l:
        section("Budget Allocation by Program ($K)")
        for _, row in hist.iterrows():
            prog  = row["Marketing Program"]
            cpa   = row["CPA"]
            cur_m = round(round(row["cost_sim"] / 1e6 / 0.5) * 0.5, 1)
            val_m = st.slider(
                f"{prog}  |  Hist. CPA: ${cpa:,.0f}",
                min_value=0.0,
                max_value=421.5,
                value=cur_m,
                step=0.5,
                format="$%.1fK",
                key=f"sim_{prog}",
            )
            new_budgets[prog] = (val_m * 1e6, cpa)

    with r:
        new_total = sum(v for v, _ in new_budgets.values())

        new_won = sum(
            budget / cpa
            for budget, cpa in new_budgets.values()
            if cpa > 0 and budget > 0
        )

        delta_won = new_won - cur_won
        delta_bud = new_total - total_budget

        new_cpa = (
            new_total / new_won
            if new_won > 0
            else 0
        )

        section("Impact Projection")

        c1, c2 = st.columns(2)

        kpi(
            c1,
            "Total Budget",
            f"${new_total/1e6:.1f}K",
            sub=f"{'↑' if delta_bud >= 0 else '↓'} ${abs(delta_bud)/1e6:.1f}K",
            color=C["amber"]
        )

        kpi(
            c2,
            "Projected Won",
            f"{new_won:,.0f}",
            sub=f"{'↑' if delta_won >= 0 else '↓'} {abs(delta_won):,.0f} vs current",
            color=C["green"] if delta_won >= 0 else C["red"]
        )

        st.markdown("<br>", unsafe_allow_html=True)

        c3, c4 = st.columns(2)

        hist_cpa = total_budget / cur_won if cur_won else 0

        kpi(
            c3,
            "Projected CPA",
            f"${new_cpa:,.0f}",
            color=C["green"] if new_cpa <= hist_cpa else C["red"]
        )

        roi_change = (
            (new_won / cur_won - 1) * 100
            if cur_won > 0
            else 0
        )

        kpi(
            c4,
            "ROI vs Current",
            f"{roi_change:+.1f}%",
            color=C["green"] if new_won >= cur_won else C["red"]
        )

        # Pareto Curve
        section("Pareto Curve — Cumulative Program Efficiency")

        par = hist.sort_values("CPA").reset_index(drop=True)

        par["cum_cost"] = par["cost_sim"].cumsum() / 1e6
        par["cum_won"] = par["won"].cumsum()

        fig = go.Figure(
            go.Scatter(
                x=par["cum_cost"],
                y=par["cum_won"],
                mode="lines+markers",
                line=dict(color=C["blue"], width=2),
                marker=dict(size=9, color=C["navy"]),
                text=par["Marketing Program"],
                hovertemplate=(
                    "<b>%{text}</b><br>"
                    "Cum. spend: $%{x:.1f}M<br>"
                    "Cum. won: %{y:,.0f}"
                    "<extra></extra>"
                ),
            )
        )

        for i, row in par.iterrows():
            fig.add_annotation(
                x=row["cum_cost"],
                y=row["cum_won"],
                text=row["Marketing Program"].replace("paid_", ""),
                showarrow=True,
                arrowhead=0,
                arrowcolor="#d1d5db",
                arrowwidth=1,
                ax=0,
                ay=-28 - (i % 3) * 18,
                font=dict(size=9, color=C["navy"]),
                bgcolor="white",
                borderpad=2,
            )

        theme(fig, h=270, legend=False)

        fig.update_xaxes(title_text="Cumulative Investment ($K)")
        fig.update_yaxes(title_text="Cumulative Won")

        st.plotly_chart(fig, use_container_width=True)
# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    st.markdown(
        """<style>
        .stApp { background-color: #f8f9fa; }
        .stTabs [data-baseweb="tab-list"] {
            gap: 4px; background: white; padding: 8px;
            border-radius: 10px; box-shadow: 0 1px 3px rgba(0,0,0,.08);
        }
        .stTabs [data-baseweb="tab"] {
            height: 40px; padding: 0 18px; border-radius: 8px;
            font-weight: 500; font-size: 13px; color: #6b7280;
        }
        .stTabs [aria-selected="true"] {
            background-color: #1e3a5f !important; color: white !important;
        }
        .block-container { padding-top: 1.5rem; max-width: 1400px; }
        </style>""",
        unsafe_allow_html=True,
    )

    df = load_data()

    # ── Sidebar filters ───────────────────────────────────────────────────────
    with st.sidebar:
        st.markdown(
            f"<div style='font-size:16px;font-weight:700;color:{C['navy']};"
            f"margin-bottom:16px'>Filters</div>",
            unsafe_allow_html=True,
        )

        all_years    = sorted(df["Year"].unique().tolist())
        all_teams    = sorted(df["Marketing Team"].unique().tolist())
        all_programs = sorted(df["Marketing Program"].unique().tolist())

        sel_years = st.multiselect("Year", all_years, default=all_years)
        sel_months = st.multiselect(
            "Month", MONTHS, default=MONTHS,
        )
        sel_teams    = st.multiselect("Marketing Team",    all_teams,    default=all_teams)
        sel_programs = st.multiselect("Marketing Program", all_programs, default=all_programs)

        sel_month_nums = [i + 1 for i, m in enumerate(MONTHS) if m in sel_months]

        if st.button("Reset filters", use_container_width=True):
            sel_years    = all_years
            sel_months   = MONTHS
            sel_teams    = all_teams
            sel_programs = all_programs
            sel_month_nums = list(range(1, 13))

    mask = (
        df["Year"].isin(sel_years)
        & df["MonthNum"].isin(sel_month_nums)
        & df["Marketing Team"].isin(sel_teams)
        & df["Marketing Program"].isin(sel_programs)
    )
    df = df[mask]

    tc = df["Cost"].sum()

    st.markdown(
        f"""<div style="background:{C['navy']};padding:20px 28px;border-radius:12px;
        margin-bottom:20px;display:flex;align-items:center;
        justify-content:space-between">
        <div>
            <div style="color:#93c5fd;font-size:11px;text-transform:uppercase;
            letter-spacing:1px">Deel · Marketing Intelligence</div>
            <div style="color:white;font-size:24px;font-weight:700;margin-top:4px">
            Marketing Analytics Dashboard</div>
            <div style="color:#94a3b8;font-size:12px;margin-top:2px">
            Jul 2023 – Jun 2025 &nbsp;·&nbsp; 11,952 records &nbsp;·&nbsp;
            7 teams &nbsp;·&nbsp; 17 programs</div>
        </div>
        <div style="text-align:right;color:#94a3b8;font-size:12px">
            Total Investment<br>
            <span style="color:white;font-size:22px;font-weight:700">${tc/1e3:.1f}K</span>
        </div>
        </div>""",
        unsafe_allow_html=True,
    )

    tabs = st.tabs([
        "Executive View",
        "Full Funnel",
        "Paid Channels",
        "Zero-Cost Channels",
        "Trends",
        "Simulator",
    ])
    with tabs[0]: p1_executive(df)
    with tabs[1]: p2_funnel(df)
    with tabs[2]: p3_paid(df)
    with tabs[3]: p4_organic(df)
    with tabs[4]: p5_trends(df)
    with tabs[5]: p6_simulator(df)


if __name__ == "__main__":
    main()
