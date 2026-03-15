# =============================================================================
# CASHIFY BRAND STUDY — DECISION SUPPORT SYSTEM (DSS)
# Dash Web App — runs locally in VS Code, opens in your browser
# =============================================================================
# SETUP (run once in terminal):
#   pip install dash dash-bootstrap-components pandas plotly openpyxl
#
# USAGE:
#   1. Place this file in the SAME FOLDER as both Excel files
#   2. In VS Code terminal: python cashify_dss.py
#   3. Open browser at: http://127.0.0.1:8050
# =============================================================================

import os
import glob
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots

import dash
from dash import dcc, html, Input, Output, callback
import dash_bootstrap_components as dbc

# ── 1. LOAD DATA ──────────────────────────────────────────────────────────────
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

def find_file(keyword):
    pattern = os.path.join(SCRIPT_DIR, f"*{keyword}*.xlsx")
    matches = glob.glob(pattern, recursive=False)
    if not matches:
        raise FileNotFoundError(
            f"Could not find an Excel file matching '*{keyword}*.xlsx' "
            f"in {SCRIPT_DIR}\n"
            f"Make sure both Excel files are in the same folder as this script."
        )
    return matches[0]

def load(keyword):
    path = find_file(keyword)
    raw  = pd.read_excel(path)
    df   = raw.iloc[1:].reset_index(drop=True)
    df.columns = raw.columns
    df["Gender"] = df["Q2"].astype(str).str.strip()
    df["Age"]    = df["Q3"].astype(str).str.strip()
    df["City"]   = df["Q1"].astype(str).str.strip()
    # Clean "nan" strings
    for col in ["Gender","Age","City"]:
        df[col] = df[col].replace("nan", np.nan)
    return df

print("Loading data...")
df_r = load("Refurbished")
df_b = load("Buyback")
print(f"  Refurbished: {len(df_r)} respondents")
print(f"  Buyback:     {len(df_b)} respondents")

# ── 2. BRAND & COLUMN CONFIG ──────────────────────────────────────────────────
BRANDS_R = [
    "Cashify", "Amazon Renewed", "Flipkart Reset", "OLX",
    "Local shop/used-phone market",
    "Facebook Marketplace / Instagram sellers",
    "Sahivalue", "XtraCover", "Refit Global", "ControlZ",
]
BRANDS_B = [
    "Cashify", "OLX", "Quikr", "Flipkart Reset", "Refit Global",
    "GetInstaCash", "PhoneCash", "Amazon Exchange", "Flipkart Exchange",
]

def col_map(brands, prefix):
    return {b: f"{prefix}_{i+1}" for i, b in enumerate(brands)}

NPS_R = col_map(BRANDS_R, "Q16")
NPS_B = col_map(BRANDS_B, "Q16")
FAM_R = col_map(BRANDS_R, "Q14")
FAM_B = col_map(BRANDS_B, "Q14")
CON_R = col_map(BRANDS_R, "Q15")
CON_B = col_map(BRANDS_B, "Q15")
SOA_R = col_map(BRANDS_R, "Q13")
SOA_B = col_map(BRANDS_B, "Q13")

CASHIFY_DRV_R = {
    "EMI available":            "Q20_1",
    "Multiple payment options": "Q20_2",
    "Verified & secure":        "Q20_3",
    "Trusted brand":            "Q20_4",
    "Recommended by friends":   "Q20_5",
    "Wide store presence":      "Q20_6",
    "32 point check":           "Q20_7",
    "Warranty included":        "Q20_8",
    "Safer than local sellers": "Q20_9",
    "Better value for money":   "Q20_10",
    "Hassle-free returns":      "Q20_11",
    "Good past experience":     "Q20_12",
    "Friends/family recommend": "Q20_13",
    "Environment-friendly":     "Q20_14",
}
COMP_DRV_R = {
    "EMI available":            "Q21A_1",
    "Multiple payment options": "Q21A_2",
    "Verified & secure":        "Q21A_3",
    "Trusted brand":            "Q21A_4",
    "Recommended by friends":   "Q21A_5",
    "Wide store presence":      "Q21A_6",
    "32 point check":           "Q21A_7",
    "Warranty included":        "Q21A_8",
    "Safer than local sellers": "Q21A_9",
    "Better value for money":   "Q21A_10",
    "Hassle-free returns":      "Q21A_11",
    "Good past experience":     "Q21A_12",
    "Friends/family recommend": "Q21A_13",
    "Environment-friendly":     "Q21A_14",
}
CASHIFY_DRV_B = {
    "Best price offered":      "Q20_1",
    "Instant payment":         "Q20_2",
    "Trusted brand":           "Q20_3",
    "Doorstep pickup":         "Q20_4",
    "Easy process":            "Q20_5",
    "Wide city coverage":      "Q20_6",
    "Good customer support":   "Q20_7",
    "Transparent pricing":     "Q20_8",
    "Referral/recommendation": "Q20_9",
    "Verified platform":       "Q20_10",
    "No negotiation hassle":   "Q20_11",
}
COMP_DRV_B = {
    "Instant payment":         "Q21A_11",
    "Best price":              "Q21A_12",
    "Technician transparency": "Q21A_13",
    "Data wiping trust":       "Q21A_14",
    "Verified & secure":       "Q21A_15",
    "Wide coverage":           "Q21A_16",
    "Customer support":        "Q21A_17",
    "App rating/reviews":      "Q21A_18",
    "Doorstep pickup":         "Q21A_19",
    "Easy process":            "Q21A_20",
    "No negotiation":          "Q21A_21",
}
SELLING_DRV_B = {
    "Best price":              "Q24_1",
    "Instant payment":         "Q24_2",
    "Safe platform/location":  "Q24_3",
    "Data wiping certificate": "Q24_4",
    "Technician behaviour":    "Q24_5",
    "Doorstep pickup":         "Q24_6",
    "Brand reputation":        "Q24_7",
    "App rating/reviews":      "Q24_8",
    "Quick process":           "Q24_9",
    "Customer support":        "Q24_10",
    "No negotiation":          "Q24_11",
    "Buyback guarantee":       "Q24_12",
    "Easy documentation":      "Q24_13",
}

# ── 3. THEME ───────────────────────────────────────────────────────────────────
BG       = "#0F172A"
CARD     = "#1E293B"
BORDER   = "#334155"
TEXT     = "#F1F5F9"
MUTED    = "#94A3B8"
CASHIFY  = "#E5432A"
PALETTE  = ["#6366F1","#22C55E","#F59E0B","#EC4899","#14B8A6",
            "#8B5CF6","#F97316","#06B6D4","#84CC16","#EF4444"]

LAYOUT = dict(
    paper_bgcolor=BG, plot_bgcolor=CARD,
    font=dict(color=TEXT, family="Inter, system-ui, sans-serif", size=12),
    margin=dict(t=55, b=50, l=60, r=30),
    xaxis=dict(gridcolor=BORDER, zerolinecolor=BORDER),
    yaxis=dict(gridcolor=BORDER, zerolinecolor=BORDER),
    legend=dict(bgcolor="rgba(0,0,0,0)", bordercolor=BORDER,
                font=dict(color=TEXT, size=11)),
    hoverlabel=dict(bgcolor=CARD, bordercolor=BORDER, font_color=TEXT),
)

def styled(fig, title="", h=480):
    fig.update_layout(**LAYOUT)
    fig.update_layout(
        height=h,
        title=dict(
            text=title,
            font=dict(size=14, color=TEXT),
            x=0.01,
            xanchor="left"
        )
    )
    return fig

# ── 4. HELPERS ────────────────────────────────────────────────────────────────
def flt(df, g, a, c):

    d = df.copy()

    if g != "All":
        d = d[d["Gender"] == g]

    if a != "All":
        d = d[d["Age"] == a]

    if c != "All":
        d = d[d["City"] == c]

    return d

def to_int(x):
    try:    return int(float(str(x).strip()))
    except: return np.nan

def calc_nps(series):
    n = series.apply(to_int).dropna()
    if len(n) == 0: return 0, 0, 0, 0
    p  = (n >= 9).sum()
    ps = ((n >= 7) & (n <= 8)).sum()
    d  = (n <= 6).sum()
    t  = len(n)
    return round(100*(p-d)/t), round(100*p/t), round(100*ps/t), round(100*d/t)

def driver_score(df, col_dict):
    scores = {}
    for label, col in col_dict.items():
        if col not in df.columns: continue
        nums = df[col].apply(to_int).dropna()
        s = nums.apply(lambda x: max(0, 6-x) if 1 <= x <= 5 else 0).sum()
        if s > 0: scores[label] = int(s)
    return scores

def explode_counts(df, col, n=12):
    if col not in df.columns: return pd.Series(dtype=int)
    return (df[col].dropna()
            .str.split(",").explode()
            .str.strip()
            .value_counts()
            .head(n))

def dropdown_opts(col, df):
    v = sorted(df[col].dropna().astype(str).unique().tolist())
    return [{"label": "All", "value": "All"}] + \
           [{"label": x, "value": x} for x in v if x]

# ── 5. CHART FUNCTIONS ────────────────────────────────────────────────────────

def chart_awareness(df, brands, lbl, g, a, c):
    d = flt(df, g, a, c); n = len(d)
    rows = []
    for brand in brands:
        bl  = brand.lower()
        tom = d["Q10"].fillna("").str.lower().str.contains(bl, regex=False).mean()*100
        sp  = d["Q11"].fillna("").str.lower().str.contains(bl, regex=False).mean()*100
        aid = d["Q12"].fillna("").str.lower().str.contains(bl, regex=False).mean()*100
        rows.append({"Brand": brand,
                     "TOM":         round(tom, 1),
                     "Spontaneous": round(min(tom+sp, 100), 1),
                     "Aided":       round(aid, 1)})
    res = pd.DataFrame(rows).sort_values("Aided", ascending=True)
    fig = go.Figure()
    for metric, color in [("Aided","#3B82F6"),("Spontaneous","#F59E0B"),("TOM",CASHIFY)]:
        fig.add_trace(go.Bar(
            y=res["Brand"], x=res[metric], name=metric, orientation="h",
            marker_color=[CASHIFY if b=="Cashify" else color for b in res["Brand"]],
            text=[f"{v}%" for v in res[metric]], textposition="auto", opacity=0.9
        ))
    fig.update_layout(barmode="overlay")
    return styled(fig, f"A.  Brand Awareness Funnel — {lbl}  (n={n})", 520)


def chart_health(df, brands, fam_cols, lbl, g, a, c):
    d = flt(df, g, a, c); n = len(d)
    NEVER = "never heard"
    EVER  = ["have used it to buy","have used it to sell","i used it","have used it recently"]
    REC   = ["recently","last 6 months","last 3 months"]
    rows  = []
    for brand in brands:
        bl  = brand.lower()
        col = fam_cols.get(brand, "")
        aw  = d["Q12"].fillna("").str.lower().str.contains(bl, regex=False).mean()*100
        if col and col in d.columns:
            v  = d[col].fillna("").str.lower()
            fm = v.apply(lambda x: NEVER not in x and x != "").mean()*100
            eu = v.apply(lambda x: any(k in x for k in EVER)).mean()*100
            ru = v.apply(lambda x: any(k in x for k in REC)).mean()*100
        else:
            fm = eu = ru = 0.0
        rows.append({
            "Brand":       brand,
            "Awareness":   round(aw, 1),
            "Familiarity": round(fm, 1),
            "Ever Used":   round(eu, 1),
            "Recent Use":  round(ru, 1),
            "Aw→Fm":  round(fm/aw*100, 1) if aw > 0 else 0,
            "Fm→Eu":  round(eu/fm*100, 1) if fm > 0 else 0,
            "Eu→Ru":  round(ru/eu*100, 1) if eu > 0 else 0,
        })
    res = pd.DataFrame(rows)
    fig = make_subplots(rows=2, cols=1, row_heights=[0.65, 0.35],
                        subplot_titles=["Brand Health Stages (%)",
                                        "Stage-to-Stage Conversion Rates (%)"])
    stage_colors = [("Awareness","#6366F1"),("Familiarity","#8B5CF6"),
                    ("Ever Used","#EC4899"),("Recent Use","#F97316")]
    for s, col in stage_colors:
        fig.add_trace(go.Bar(name=s, x=res["Brand"], y=res[s],
                             marker_color=col,
                             text=[f"{v}%" for v in res[s]],
                             textposition="outside"), row=1, col=1)
    for cv, cl, col in [("Aw→Fm","Aware→Familiar","#A78BFA"),
                         ("Fm→Eu","Familiar→Used","#F472B6"),
                         ("Eu→Ru","Used→Recent","#FB923C")]:
        fig.add_trace(go.Bar(name=cl, x=res["Brand"], y=res[cv],
                             marker_color=col,
                             text=[f"{v}%" for v in res[cv]],
                             textposition="outside"), row=2, col=1)
    fig.update_layout(barmode="group", **LAYOUT,
                      title=dict(text=f"B.  Brand Health Funnel + Conversion Rates — {lbl}  (n={n})",
                                 font=dict(size=14, color=TEXT)),
                      height=680,
                      xaxis2=dict(gridcolor=BORDER, zerolinecolor=BORDER),
                      yaxis2=dict(gridcolor=BORDER, zerolinecolor=BORDER))
    return fig


def chart_nps(df, brands, nps_cols, lbl, g, a, c):
    d = flt(df, g, a, c); n = len(d)
    rows = []
    for brand in brands:
        col = nps_cols.get(brand, "")
        if not col or col not in d.columns: continue
        nps, pro, pas, det = calc_nps(d[col])
        rows.append({"Brand": brand, "NPS": nps,
                     "Promoters": pro, "Passives": pas, "Detractors": det})
    if not rows: return go.Figure()
    res = pd.DataFrame(rows).sort_values("NPS", ascending=False)
    bar_col = [CASHIFY if b == "Cashify" else ("#22C55E" if v >= 0 else "#EF4444")
               for b, v in zip(res["Brand"], res["NPS"])]
    fig = make_subplots(rows=1, cols=2, column_widths=[0.42, 0.58],
                        subplot_titles=["NPS Score by Brand",
                                        "Promoter / Passive / Detractor %"])
    fig.add_trace(go.Bar(y=res["Brand"], x=res["NPS"], orientation="h",
                         marker_color=bar_col, name="NPS",
                         text=res["NPS"], textposition="outside"), row=1, col=1)
    for m, col in [("Promoters","#22C55E"),("Passives","#F59E0B"),("Detractors","#EF4444")]:
        fig.add_trace(go.Bar(name=m, x=res["Brand"], y=res[m], marker_color=col,
                             text=[f"{v}%" for v in res[m]],
                             textposition="inside"), row=1, col=2)
    fig.update_layout(barmode="stack", **LAYOUT,
                      title=dict(text=f"C.  NPS Dashboard — {lbl}  (n={n})",
                                 font=dict(size=14, color=TEXT)),
                      height=520,
                      xaxis2=dict(gridcolor=BORDER, zerolinecolor=BORDER),
                      yaxis2=dict(gridcolor=BORDER, zerolinecolor=BORDER))
    return fig

def chart_brand_funnel(df, brands, fam_cols, con_cols, lbl, g, a, c):

    d = flt(df, g, a, c)
    n = len(d)

    rows = []

    for brand in brands:

        aware = d["Q12"].fillna("").str.lower().str.contains(brand.lower(), regex=False).mean()*100

        fam_col = fam_cols.get(brand,"")
        if fam_col in d.columns:
            fam = d[fam_col].notna().mean()*100
        else:
            fam = 0

        con_col = con_cols.get(brand,"")
        if con_col in d.columns:
            consider = d[con_col].str.contains("consider",case=False,na=False).mean()*100
        else:
            consider = 0

        rows.append({
            "Brand":brand,
            "Awareness":round(aware,1),
            "Familiarity":round(fam,1),
            "Consideration":round(consider,1)
        })

    res = pd.DataFrame(rows)

    stages = ["Awareness","Familiarity","Consideration"]

    fig = go.Figure()

    for i,row in res.iterrows():

        fig.add_trace(go.Scatter(
            x=stages,
            y=[row["Awareness"],row["Familiarity"],row["Consideration"]],
            mode="lines+markers",
            name=row["Brand"],
            line=dict(width=4 if row["Brand"]=="Cashify" else 2,
                      color=CASHIFY if row["Brand"]=="Cashify" else PALETTE[i%len(PALETTE)])
        ))

    fig.update_layout(
        **LAYOUT,
        height=500,
        title=f"I. Brand Funnel (McKinsey Style) — {lbl} (n={n})"
    )

    return fig
def chart_positioning(df, brands, nps_cols, fam_cols, con_cols, lbl, g, a, c):

    d = flt(df,g,a,c)
    n = len(d)

    rows=[]

    for brand in brands:

        aware = d["Q12"].fillna("").str.lower().str.contains(brand.lower(),regex=False).mean()*100

        con_col = con_cols.get(brand,"")
        if con_col in d.columns:
            consider = d[con_col].str.contains("consider",case=False,na=False).mean()*100
        else:
            consider=0

        nps_col = nps_cols.get(brand,"")
        if nps_col in d.columns:
            nps,_p,_pa,_d = calc_nps(d[nps_col])
        else:
            nps=0

        rows.append({
            "Brand":brand,
            "Awareness":aware,
            "Consideration":consider,
            "NPS":nps
        })

    res = pd.DataFrame(rows)

    fig = px.scatter(
        res,
        x="Awareness",
        y="NPS",
        size="Consideration",
        text="Brand"
    )

    fig.update_traces(
        marker=dict(
            color=[CASHIFY if b=="Cashify" else "#6366F1" for b in res["Brand"]],
            line=dict(width=1,color="#000")
        )
    )

    fig.update_layout(
        **LAYOUT,
        height=520,
        title=f"J. Market Positioning Map — {lbl} (n={n})"
    )

    fig.update_xaxes(title="Brand Awareness (%)")
    fig.update_yaxes(title="Net Promoter Score")

    return fig

def chart_soa(df, brands, soa_cols, lbl, g, a, c):
    """
    Source of Awareness: grouped horizontal bar chart.
    For each CHANNEL (y-axis), shows % of respondents who cited it
    for each BRAND (separate coloured bars).
    Completely distinct from the Consideration Set stacked bars.
    """
    d = flt(df, g, a, c); n = len(d)

    # ── Collect all unique channels ──────────────────────────────────────────
    all_ch = set()
    for brand in brands:
        col = soa_cols.get(brand, "")
        if col and col in d.columns:
            d[col].dropna().str.split(",").explode().str.strip().apply(
                lambda x: all_ch.add(x) if x and x.lower() not in ("nan", "") else None)
    channels = sorted([ch for ch in all_ch if ch and len(ch) > 2])
    if not channels:
        fig = go.Figure()
        fig.add_annotation(text="No Source of Awareness data for this filter",
                           xref="paper", yref="paper", x=0.5, y=0.5,
                           showarrow=False, font=dict(size=14, color=MUTED))
        fig.update_layout(**LAYOUT, height=400)
        fig.update_layout(title_text="D.  Source of Awareness — No data")
        return fig

    # Shorten channel labels for display
    ch_labels = [ch.replace("Saw as brand integration on a Youtube show",
                             "YouTube brand integration")
                   .replace("Price–comparison site/Deal–blog", "Price comparison / Deal blog")
                   .replace("Social media ad (Instagram/Facebook)", "Social media ad")
                   .replace("Influencer video recommending it", "Influencer video")
                   .replace("YouTube review/Unboxing", "YouTube review/Unboxing")
                 for ch in channels]

    # ── Build % matrix: rows=brands, cols=channels ───────────────────────────
    # Denominator = respondents aware of that brand (non-null Q13 col)
    # This gives "of those aware of Brand X, what % heard via each channel"
    brand_colors = {}
    for i, brand in enumerate(brands):
        brand_colors[brand] = CASHIFY if brand == "Cashify" else PALETTE[i % len(PALETTE)]

    fig = go.Figure()
    for brand in brands:
        col = soa_cols.get(brand, "")
        if not col or col not in d.columns:
            continue
        base = d[col].notna().sum()          # only aware respondents as denominator
        if base == 0:
            continue
        pcts = []
        for ch in channels:
            cnt = d[col].fillna("").str.lower()\
                         .str.contains(ch.lower(), regex=False).sum()
            pcts.append(round(cnt / base * 100, 1))

        fig.add_trace(go.Bar(
            name=brand,
            y=ch_labels,
            x=pcts,
            orientation="h",
            marker_color=brand_colors.get(brand, "#6366F1"),
            text=[f"{v}%" if v > 0 else "" for v in pcts],
            textposition="outside",
            hovertemplate=(
                f"<b>{brand}</b><br>"
                "Channel: %{y}<br>"
                "% Aware audience: %{x}%<br>"
                f"(Base: {base} aware respondents)"
                "<extra></extra>"
            )
        ))

    fig.update_layout(
    **LAYOUT,
    barmode="group",
    height=max(500, 55 * len(channels) + 120),
    )
    fig.update_layout(
    margin=dict(t=80, b=50, l=220, r=160)
    )
    fig.update_layout(title_text=(
        f"D.  Source of Awareness — How consumers discovered each platform  |  "
        f"{lbl} study | n={n} | % of brand-aware respondents per channel"
    ))
    fig.update_xaxes(
        title_text="% of brand-aware respondents",
        gridcolor=BORDER, zerolinecolor=BORDER, ticksuffix="%"
    )
    fig.update_yaxes(
        gridcolor=BORDER, zerolinecolor=BORDER,
        tickfont=dict(size=11), autorange="reversed"
    )
    return fig


def chart_consideration(df, brands, con_cols, lbl, g, a, c):
    d = flt(df, g, a, c); n = len(d)
    rows = []
    for brand in brands:
        col = con_cols.get(brand, "")
        if not col or col not in d.columns: continue
        v = d[col].fillna("")
        rows.append({
            "Brand":             brand,
            "First Choice":      round(v.str.contains("first choice",  case=False).mean()*100, 1),
            "Seriously Consider":round(v.str.contains("seriously consider", case=False).mean()*100, 1),
            "Might Consider":    round(v.str.contains("might consider", case=False).mean()*100, 1),
            "Would Not Consider":round(v.str.contains("would not",      case=False).mean()*100, 1),
        })
    if not rows: return go.Figure()
    res = pd.DataFrame(rows).sort_values("First Choice", ascending=False)
    fig = go.Figure()
    for col, color in [("First Choice","#22C55E"),("Seriously Consider","#84CC16"),
                        ("Might Consider","#F59E0B"),("Would Not Consider","#EF4444")]:
        fig.add_trace(go.Bar(name=col, x=res["Brand"], y=res[col],
                             marker_color=color,
                             text=[f"{v}%" for v in res[col]],
                             textposition="inside"))
    fig.update_layout(barmode="stack")
    return styled(fig, f"E.  Consideration Set — {lbl}  (n={n})", 500)


def chart_drivers(df, cashify_cols, comp_cols, lbl, g, a, c):
    d = flt(df, g, a, c); n = len(d)

    # Q20 = only Cashify buyers answered | Q21A = only non-Cashify buyers answered
    # Determine base sizes for honest labelling
    cashify_base = int(d[[col for col in cashify_cols.values() if col in d.columns][0:1]
                         ].notna().any(axis=1).sum()) if cashify_cols else 0
    comp_base    = int(d[[col for col in comp_cols.values()    if col in d.columns][0:1]
                         ].notna().any(axis=1).sum()) if comp_cols else 0

    cs = driver_score(d, cashify_cols)
    ks = driver_score(d, comp_cols)

    # If both empty → show informative empty state
    if not cs and not ks:
        fig = go.Figure()
        fig.add_annotation(
            text=(f"No choice driver data available for this filter<br>"
                  f"<i>Choice drivers are only answered by actual buyers/sellers.<br>"
                  f"Try selecting 'All' for age/gender/city to see aggregate results.</i>"),
            xref="paper", yref="paper", x=0.5, y=0.5,
            showarrow=False, font=dict(size=14, color=MUTED),
            align="center"
        )
        return fig.update_layout(**LAYOUT, height=400,
            title=dict(text=f"F.  Choice Drivers — {lbl} (Insufficient data for selected filter)",
                       font=dict(size=14, color=TEXT)))

    all_f = sorted(set(list(cs) + list(ks)))
    mx    = max(list(cs.values()) + list(ks.values()) + [1])
    c_n   = [round(cs.get(f, 0)/mx*100, 1) for f in all_f]
    k_n   = [round(ks.get(f, 0)/mx*100, 1) for f in all_f]
    order = sorted(range(len(all_f)), key=lambda i: c_n[i])
    fsort = [all_f[i] for i in order]
    csort = [c_n[i]   for i in order]
    ksort = [k_n[i]   for i in order]

    cashify_label = f"Cashify buyers (n={cashify_base})" if cashify_base > 0 else "Cashify (no data)"
    comp_label    = f"Other platforms (n={comp_base})"   if comp_base   > 0 else "Competitors (no data)"

    fig = go.Figure()
    if any(v > 0 for v in csort):
        fig.add_trace(go.Bar(
            y=fsort, x=csort, name=cashify_label, orientation="h",
            marker_color=CASHIFY,
            text=[f"{v}" if v > 0 else "" for v in csort],
            textposition="outside"
        ))
    if any(v > 0 for v in ksort):
        fig.add_trace(go.Bar(
            y=fsort, x=ksort, name=comp_label, orientation="h",
            marker_color="#6366F1",
            text=[f"{v}" if v > 0 else "" for v in ksort],
            textposition="outside"
        ))

    # Add base size note
    note = (f"Cashify buyers: n={cashify_base} | "
            f"Other platform buyers: n={comp_base} | "
            f"Note: Only actual buyers answer choice drivers")
    fig.update_layout(barmode="group")
    return styled(fig,
        f"F.  Choice Drivers: Cashify vs Competitors — {lbl}  (Total n={n})<br>"
        f"<sup>{note}</sup>",
        600)


def chart_barriers(df, col, lbl, g, a, c):
    d = flt(df, g, a, c); n = len(d)
    if col not in d.columns: return go.Figure()
    counts = explode_counts(d, col, 15)
    counts = counts[counts.index.str.len() > 3].sort_values(ascending=True)
    fig = go.Figure(go.Bar(y=counts.index, x=counts.values, orientation="h",
                           marker_color="#EF4444",
                           text=counts.values, textposition="outside"))
    return styled(fig, f"G.  Barriers to Choosing Cashify — {lbl}  (n={n})", 520)

def chart_radar(df, brands, fam_cols, con_cols, nps_cols, lbl, g, a, c):

    d = flt(df, g, a, c)

    rows = []

    for brand in brands:

        awareness = d["Q12"].fillna("").str.lower().str.contains(
            brand.lower(), regex=False).mean()*100

        fam_col = fam_cols.get(brand,"")
        if fam_col in d.columns:
            familiarity = d[fam_col].notna().mean()*100
        else:
            familiarity = 0

        con_col = con_cols.get(brand,"")
        if con_col in d.columns:
            consideration = d[con_col].str.contains(
                "consider",case=False,na=False).mean()*100
        else:
            consideration = 0

        nps_col = nps_cols.get(brand,"")
        if nps_col in d.columns:
            nps,_,_,_ = calc_nps(d[nps_col])
        else:
            nps = 0

        rows.append({
            "Brand":brand,
            "Awareness":awareness,
            "Familiarity":familiarity,
            "Consideration":consideration,
            "NPS":max(0,nps+50)  # normalize for radar
        })

    res = pd.DataFrame(rows)

    categories = ["Awareness","Familiarity","Consideration","NPS"]

    fig = go.Figure()

    for i,row in res.iterrows():

        values = [
            row["Awareness"],
            row["Familiarity"],
            row["Consideration"],
            row["NPS"]
        ]

        values.append(values[0])

        fig.add_trace(go.Scatterpolar(
            r=values,
            theta=categories+ [categories[0]],
            fill='toself',
            name=row["Brand"],
            line=dict(
                color=CASHIFY if row["Brand"]=="Cashify"
                else PALETTE[i % len(PALETTE)]
            )
        ))

    fig.update_layout(
        polar=dict(radialaxis=dict(visible=True,range=[0,100])),
        showlegend=True,
        **LAYOUT,
        height=520,
        title=f"K. Competitive Brand Radar — {lbl}"
    )

    return fig
def chart_brand_power(df, brands, fam_cols, con_cols, nps_cols, lbl, g, a, c):

    d = flt(df, g, a, c)

    rows = []

    for brand in brands:

        awareness = d["Q12"].fillna("").str.lower().str.contains(
            brand.lower(),regex=False).mean()*100

        fam_col = fam_cols.get(brand,"")
        familiarity = d[fam_col].notna().mean()*100 if fam_col in d.columns else 0

        con_col = con_cols.get(brand,"")
        consideration = d[con_col].str.contains(
            "consider",case=False,na=False).mean()*100 if con_col in d.columns else 0

        nps_col = nps_cols.get(brand,"")
        if nps_col in d.columns:
            nps,_,_,_ = calc_nps(d[nps_col])
        else:
            nps = 0

        score = (
            0.40*awareness +
            0.25*familiarity +
            0.20*consideration +
            0.15*(nps+50)
        )

        rows.append({
            "Brand":brand,
            "Power Score":round(score,1)
        })

    res = pd.DataFrame(rows).sort_values("Power Score",ascending=True)

    fig = go.Figure(go.Bar(
        y=res["Brand"],
        x=res["Power Score"],
        orientation="h",
        marker_color=[
            CASHIFY if b=="Cashify" else "#6366F1"
            for b in res["Brand"]
        ],
        text=res["Power Score"],
        textposition="outside"
    ))

    fig.update_layout(
        **LAYOUT,
        height=500,
        title=f"L. Brand Power Index — {lbl}"
    )

    return fig

def chart_category(df, is_refurb, lbl, g, a, c):
    d = flt(df, g, a, c); n = len(d)
    if is_refurb:
        drv  = explode_counts(d, "Q22", 10)
        fear = explode_counts(d, "Q23", 10)
        fig  = make_subplots(rows=1, cols=2,
                             subplot_titles=["Category Purchase Drivers (Q22)",
                                             "Category Fears (Q23)"])
        if not drv.empty:
            fig.add_trace(go.Bar(y=drv.index[::-1], x=drv.values[::-1],
                                 orientation="h", marker_color="#22C55E",
                                 name="Drivers",
                                 text=drv.values[::-1], textposition="outside"),
                          row=1, col=1)
        if not fear.empty:
            fig.add_trace(go.Bar(y=fear.index[::-1], x=fear.values[::-1],
                                 orientation="h", marker_color="#EF4444",
                                 name="Fears",
                                 text=fear.values[::-1], textposition="outside"),
                          row=1, col=2)
        h = 520
    else:
        trust = explode_counts(d, "Q22", 10)
        fear  = explode_counts(d, "Q23", 7)
        sell_scores = {f: int(d[c].notna().sum())
                       for f, c in SELLING_DRV_B.items()
                       if c in d.columns and d[c].notna().sum() > 0}
        sell_sr = pd.Series(sell_scores).sort_values(ascending=True)
        fig = make_subplots(rows=1, cols=3,
                            subplot_titles=["Trust Factors (Q22)",
                                            "Biggest Fears (Q23)",
                                            "Selling Platform Drivers (Q24)"])
        if not trust.empty:
            fig.add_trace(go.Bar(y=trust.index[::-1], x=trust.values[::-1],
                                 orientation="h", marker_color="#22C55E",
                                 name="Trust",
                                 text=trust.values[::-1], textposition="outside"),
                          row=1, col=1)
        if not fear.empty:
            fig.add_trace(go.Bar(y=fear.index[::-1], x=fear.values[::-1],
                                 orientation="h", marker_color="#EF4444",
                                 name="Fears",
                                 text=fear.values[::-1], textposition="outside"),
                          row=1, col=2)
        if not sell_sr.empty:
            fig.add_trace(go.Bar(y=sell_sr.index, x=sell_sr.values,
                                 orientation="h", marker_color="#F59E0B",
                                 name="Sell Drivers",
                                 text=sell_sr.values, textposition="outside"),
                          row=1, col=3)
        h = 560
    fig.update_layout(**LAYOUT,
                      title=dict(text=f"H.  Category Ecosystem — {lbl}  (n={n})",
                                 font=dict(size=14, color=TEXT)),
                      height=h,
                      xaxis2=dict(gridcolor=BORDER, zerolinecolor=BORDER),
                      yaxis2=dict(gridcolor=BORDER, zerolinecolor=BORDER))
    return fig
def chart_market_opportunity(df, brands, lbl, g, a, c):

    d = flt(df, g, a, c)

    segments = ["Gender", "Age", "City"]

    rows = []

    for seg in segments:

        values = d[seg].dropna().unique()

        for v in values:

            sub = d[d[seg] == v]

            cash_aw = sub["Q12"].fillna("").str.lower().str.contains(
                "cashify", regex=False).mean()*100

            comp_aw = sub["Q12"].fillna("").str.lower().str.contains(
                "|".join([b.lower() for b in brands if b != "Cashify"]),
                regex=True).mean()*100

            opportunity = comp_aw - cash_aw

            rows.append({
                "Segment": f"{seg}: {v}",
                "Cashify Awareness": round(cash_aw,1),
                "Competitor Awareness": round(comp_aw,1),
                "Opportunity Score": round(opportunity,1)
            })

    res = pd.DataFrame(rows)

    fig = px.imshow(
        res[["Cashify Awareness","Competitor Awareness","Opportunity Score"]],
        labels=dict(x="Metric",y="Segment",color="Score"),
        y=res["Segment"],
        color_continuous_scale="RdYlGn_r"
    )

    fig.update_layout(
        **LAYOUT,
        height=520,
        title=f"M. Market Opportunity Heatmap — {lbl}"
    )

    return fig

def detect_market_opportunities(df, brands):

    insights = []

    genders = df["Gender"].dropna().unique()

    for g in genders:

        sub = df[df["Gender"] == g]

        cash_aw = sub["Q12"].fillna("").str.lower().str.contains(
            "cashify", regex=False).mean()*100

        comp_aw = sub["Q12"].fillna("").str.lower().str.contains(
            "|".join([b.lower() for b in brands if b != "Cashify"]),
            regex=True).mean()*100

        if comp_aw - cash_aw > 20:

            insights.append(
                f"{g} segment shows strong competitor awareness but low Cashify awareness — marketing opportunity."
            )

    if len(insights) == 0:

        insights.append(
            "No major awareness gaps detected across demographic segments."
        )

    return insights

def channel_strategy_insights(df):

    insights = []

    if "Q13" not in df.columns:
        return insights

    channels = (
        df["Q13"]
        .dropna()
        .str.split(",")
        .explode()
        .str.strip()
    )

    counts = channels.value_counts()

    if len(counts) == 0:
        return insights

    top_channel = counts.index[0]

    insights.append(
        f"{top_channel} is the strongest acquisition channel."
    )

    if "influencer" in top_channel.lower():
        insights.append(
            "Influencer marketing appears highly effective for discovery."
        )

    if "youtube" in top_channel.lower():
        insights.append(
            "Video content is a major driver of brand discovery."
        )

    if counts.iloc[0] > counts.mean():
        insights.append(
            "Channel concentration is high — diversifying acquisition channels may reduce dependency."
        )

    return insights

def chart_channel_attribution(df, brands, soa_cols, lbl, g, a, c):

    d = flt(df, g, a, c)

    all_channels = []

    for brand in brands:

        col = soa_cols.get(brand, "")

        if col in d.columns:

            channels = (
                d[col]
                .dropna()
                .str.split(",")
                .explode()
                .str.strip()
            )

            all_channels.extend(channels.tolist())

    if len(all_channels) == 0:

        fig = go.Figure()

        fig.add_annotation(
            text="No channel data available",
            x=0.5,
            y=0.5,
            showarrow=False,
            font=dict(size=16)
        )

        fig.update_layout(**LAYOUT)

        return fig

    counts = pd.Series(all_channels).value_counts()

    res = pd.DataFrame({
        "Channel": counts.index,
        "Count": counts.values
    })

    res["Awareness %"] = res["Count"] / res["Count"].sum() * 100

    fig = px.bar(
        res,
        x="Awareness %",
        y="Channel",
        orientation="h",
        text=res["Awareness %"].round(1).astype(str) + "%",
        color="Awareness %",
        color_continuous_scale="Blues"
    )

    fig.update_layout(
        **LAYOUT,
        height=520,
        title=f"N. Customer Acquisition Channel Attribution — {lbl}"
    )

    fig.update_yaxes(autorange="reversed")

    return fig

def chart_growth_simulator(df, brands, fam_cols, con_cols, nps_cols, awareness_boost, lbl, g, a, c):

    d = flt(df, g, a, c)

    rows = []

    for brand in brands:

        awareness = d["Q12"].fillna("").str.lower().str.contains(
            brand.lower(), regex=False).mean()*100

        fam_col = fam_cols.get(brand,"")
        familiarity = d[fam_col].notna().mean()*100 if fam_col in d.columns else 0

        con_col = con_cols.get(brand,"")
        consideration = d[con_col].str.contains(
            "consider",case=False,na=False).mean()*100 if con_col in d.columns else 0

        nps_col = nps_cols.get(brand,"")
        if nps_col in d.columns:
            nps,_,_,_ = calc_nps(d[nps_col])
        else:
            nps = 0

        # Apply simulated awareness increase
        new_awareness = min(100, awareness + awareness_boost)

        # Assume funnel conversion ratios remain constant
        fam_ratio = familiarity/awareness if awareness>0 else 0
        con_ratio = consideration/familiarity if familiarity>0 else 0

        new_fam = new_awareness * fam_ratio
        new_cons = new_fam * con_ratio

        rows.append({
            "Brand": brand,
            "Awareness": awareness,
            "Simulated Awareness": new_awareness,
            "Simulated Consideration": new_cons,
            "NPS": nps
        })

    res = pd.DataFrame(rows)

    fig = px.bar(
        res,
        x="Brand",
        y=["Awareness","Simulated Awareness"],
        barmode="group",
        text_auto=True
    )

    fig.update_layout(
        **LAYOUT,
        height=520,
        title=f"O. Brand Growth Simulator (+{awareness_boost}% Awareness) — {lbl}"
    )

    return fig

# ── 6. DASH LAYOUT ────────────────────────────────────────────────────────────
app = dash.Dash(
    __name__,
    external_stylesheets=[dbc.themes.SLATE],
    title="Cashify DSS",
)

app.index_string = """<!DOCTYPE html>
<html>
<head>
{%metas%}
<title>{%title%}</title>
{%favicon%}
{%css%}
<style>

  /* Dropdown control box */
  .Select-control {
    background-color: #FFFFFF !important;
    border-color: #CBD5E1 !important;
    color: #000000 !important;
  }

  /* Typed text / selected value */
  .Select-value-label,
  .Select--single > .Select-control .Select-value,
  .Select-placeholder,
  .Select-input > input {
    color: #000000 !important;
  }

  /* Dropdown menu panel */
  .Select-menu-outer {
    background-color: #1E293B !important;
    border-color: #334155 !important;
    z-index: 9999 !important;
  }

  /* Each option row */
  .Select-option {
    background-color: #1E293B !important;
    color: #F1F5F9 !important;
  }

  .Select-option:hover,
  .Select-option.is-focused {
    background-color: #334155 !important;
    color: #F1F5F9 !important;
  }

  .Select-option.is-selected {
    background-color: #E5432A !important;
    color: #FFFFFF !important;
  }

  /* Arrow icon */
  .Select-arrow { border-top-color: #94A3B8 !important; }
  .is-open .Select-arrow { border-bottom-color: #94A3B8 !important; }

  /* Clear X button */
  .Select-clear { color: #94A3B8 !important; }

  /* Page background */
  body { background-color: #0F172A !important; }

  /* ========================= */
  /* Slider styling */
  /* ========================= */

  /* Slider tick labels (0%,10%,20%,...) */
  .rc-slider-mark-text,
  .rc-slider-mark-text-active,
  .rc-slider-mark span {
      color: #FFFFFF !important;
      font-size: 12px !important;
      font-weight: 500 !important;
  }

  /* Active purple bar */
  .rc-slider-track {
      background-color: #A78BFA !important;
  }

  /* Background rail */
  .rc-slider-rail {
      background-color: #334155 !important;
  }

  /* Slider handle */
  .rc-slider-handle {
      border: 3px solid #A78BFA !important;
      background-color: #FFFFFF !important;
  }

  /* Tooltip box (if shown) */
  .rc-slider-tooltip-inner {
      color: #000000 !important;
      background-color: #FFFFFF !important;
      font-weight: 700 !important;
      font-size: 14px !important;
  }

  /* Tooltip arrow */
  .rc-slider-tooltip-arrow {
      border-top-color: #FFFFFF !important;
  }

</style>
</head>
<body>
{%app_entry%}
<footer>
{%config%}
{%scripts%}
{%renderer%}
</footer>
</body>
</html>
"""

# Shared dropdown style
DD = {"backgroundColor": "#FFFFFF", "color": "#000000", "border": "1px solid #CBD5E1"}
DD_LABEL = {"color": MUTED, "fontSize": "11px",
            "textTransform": "uppercase", "letterSpacing": "0.05em",
            "marginBottom": "4px"}

def kpi_card(label, value, color=CASHIFY, sub=""):
    return dbc.Col(
        html.Div([
            html.Div(label, style={**DD_LABEL}),
            html.Div(value, style={"fontSize": "28px", "fontWeight": "800",
                                   "color": color, "lineHeight": "1.1"}),
            html.Div(sub,   style={"fontSize": "11px", "color": "#64748B"}),
        ], style={"background": CARD, "borderRadius": "10px",
                  "padding": "14px 18px", "border": f"1px solid {BORDER}"}),
        style={"padding": "4px"}
    )

SECTION_BUTTONS = [
    ("A", "Awareness Funnel"),
    ("B", "Brand Health"),
    ("C", "NPS"),
    ("D", "Source of Awareness"),
    ("E", "Consideration Set"),
    ("F", "Choice Drivers"),
    ("G", "Barriers"),
    ("H", "Category Ecosystem"),
    ("I", "Brand Funnel"),
    ("J", "Market Positioning Map"),
    ("K","Competitive Radar"),
    ("L","Brand Power Index"),
    ("M","Market Opportunity Heatmap"),
    ("N","Channel Attribution"),
    ("O","Brand Growth Simulator")
]

app.layout = html.Div(style={"backgroundColor": BG, "minHeight": "100vh",
                              "fontFamily": "Inter, system-ui, sans-serif",
                              "color": TEXT, "padding": "20px"}, children=[

    # ── Header ─────────────────────────────────────────────────────────────
    html.Div(style={"background": "linear-gradient(135deg,#E5432A,#7C3AED)",
                    "borderRadius": "12px", "padding": "18px 24px",
                    "marginBottom": "20px"}, children=[
        html.H1("Cashify Brand Study — Decision Support System (DSS)",
                style={"margin": 0, "fontSize": "22px", "color": "#fff",
                       "fontWeight": "800"}),
        html.P("Refurbished (Buy) Study  |  Buyback (Sell) Study  |  15 Brand Intelligence Modules (A–O)",
               style={"margin": "4px 0 0", "fontSize": "12px",
                      "color": "rgba(255,255,255,0.8)"}),
    ]),

    # ── Filters row ────────────────────────────────────────────────────────
    dbc.Row(style={"marginBottom": "16px"}, children=[
        dbc.Col(width=2, children=[
            html.Div("Study", style=DD_LABEL),
            dcc.Dropdown(id="dd-study",
                         options=[{"label": "Refurbished (Buy)",  "value": "R"},
                                  {"label": "Buyback (Sell)", "value": "B"}],
                         value="R", clearable=False, style=DD),
        ]),
        dbc.Col(width=2, children=[
            html.Div("Gender", style=DD_LABEL),
            dcc.Dropdown(id="dd-gender", value="All", clearable=False, style=DD),
        ]),
        dbc.Col(width=2, children=[
            html.Div("Age Group", style=DD_LABEL),
            dcc.Dropdown(id="dd-age", value="All", clearable=False, style=DD),
        ]),
        dbc.Col(width=3, children=[
            html.Div("City", style=DD_LABEL),
            dcc.Dropdown(id="dd-city", value="All", clearable=False, style=DD),
        ]),
    ]),

    # ── Section selector ────────────────────────────────────────────────────
    html.Div(style={"display": "flex", "flexWrap": "wrap", "gap": "8px",
                    "marginBottom": "16px"}, children=[
        html.Button(
            f"{letter}. {name}",
            id=f"btn-{letter}",
            n_clicks=0,
            style={"background": CARD, "color": TEXT,
                   "border": f"2px solid {BORDER}",
                   "borderRadius": "8px", "padding": "8px 16px",
                   "cursor": "pointer", "fontSize": "13px",
                   "fontFamily": "inherit", "fontWeight": "500",
                   "transition": "all 0.15s ease"}
        ) for letter, name in SECTION_BUTTONS
    ]),

    # Active section indicator
    html.Div(id="section-indicator",
             style={"background": CARD, "borderLeft": f"4px solid {CASHIFY}",
                    "padding": "8px 14px", "borderRadius": "6px",
                    "fontSize": "13px", "marginBottom": "12px",
                    "color": TEXT}),

# ── KPI strip ──────────────────────────────────────────────────────────
dbc.Row(id="kpi-strip", style={"marginBottom": "12px"}),

# ── Growth simulator control ───────────────────────────────────────────
html.Div(
    [
        html.Div("Awareness Growth Simulation", style=DD_LABEL),

        dcc.Slider(
            id="awareness-slider",
            min=0,
            max=40,
            step=5,
            value=10,
            marks={0:"0%",10:"10%",20:"20%",30:"30%",40:"40%"},
        ),

        html.Div(
    id="slider-value",
    style={
        "marginTop": "6px",
        "color": "#000000",
        "background": "#FFFFFF",
        "padding": "6px 10px",
        "borderRadius": "6px",
        "display": "inline-block",
        "fontWeight": "600"
    }
)
    ],
    id="simulator-controls",
    style={"marginBottom":"15px"},
),

# ── Chart ─────────────────────────────────────────────────────────────
dcc.Graph(
    id="main-chart",
    config={"displayModeBar": True, "responsive": True},
    style={"borderRadius": "10px", "overflow": "hidden"},
),

    # Hidden store for active section
    dcc.Store(id="active-section", data="A"),
])

# ── 7. CALLBACKS ──────────────────────────────────────────────────────────────

# Update filter options when study changes
@app.callback(
        
    Output("dd-gender", "options"),
    Output("dd-age", "options"),
    Output("dd-city", "options"),
    Output("dd-gender", "value"),
    Output("dd-age", "value"),
    Output("dd-city", "value"),
    Input("dd-study", "value"),
    
)

def update_filter_options(study):
    df = df_r if study == "R" else df_b
    return (
        dropdown_opts("Gender", df),
        dropdown_opts("Age", df),
        dropdown_opts("City", df),
        "All",
        "All",
        "All",
    )


# Capture which section button was last clicked
@app.callback(
    Output("active-section", "data"),
    [Input(f"btn-{l}", "n_clicks") for l, _ in SECTION_BUTTONS],
    prevent_initial_call=True,
)
def set_section(*args):
    ctx = dash.callback_context
    if not ctx.triggered:
        return "A"
    btn_id = ctx.triggered[0]["prop_id"].split(".")[0]
    return btn_id.replace("btn-", "")


# Highlight active section button
@app.callback(
    [Output(f"btn-{l}", "style") for l, _ in SECTION_BUTTONS],
    Input("active-section", "data"),
)
def highlight_button(active):

    styles = []

    for letter, _ in SECTION_BUTTONS:

        if letter == active:
            style = {
                "background": CASHIFY,
                "color": "#fff",
                "border": f"2px solid {CASHIFY}",
                "borderRadius": "8px",
                "padding": "8px 16px",
                "cursor": "pointer",
                "fontSize": "13px",
                "fontWeight": "700",
            }
        else:
            style = {
                "background": CARD,
                "color": TEXT,
                "border": f"2px solid {BORDER}",
                "borderRadius": "8px",
                "padding": "8px 16px",
                "cursor": "pointer",
                "fontSize": "13px",
            }

        styles.append(style)

    return styles


@app.callback(
    Output("slider-value","children"),
    Input("awareness-slider","value")
)
def show_slider_value(v):
    return f"Simulated awareness increase: +{v}%"
def show_slider_value(v):
    return f"Simulated awareness increase: +{v}%"
def highlight_button(active):

    styles = []

    for letter, _ in SECTION_BUTTONS:

        if letter == active:
            style = {
                "background": CASHIFY,
                "color": "#fff",
                "border": f"2px solid {CASHIFY}",
                "borderRadius": "8px",
                "padding": "8px 16px",
                "cursor": "pointer",
                "fontSize": "13px",
                "fontWeight": "700",
            }
        else:
            style = {
                "background": CARD,
                "color": TEXT,
                "border": f"2px solid {BORDER}",
                "borderRadius": "8px",
                "padding": "8px 16px",
                "cursor": "pointer",
                "fontSize": "13px",
            }

        styles.append(style)

    return styles

@app.callback(
    Output("simulator-controls","style"),
    Input("active-section","data")
)
def toggle_simulator_controls(section):

    if section == "O":
        return {"marginBottom":"15px","display":"block"}

    return {"display":"none"}


# ── Insight generator ─────────────────────────────────────────────────────────

def generate_insights(df, brands):

    insights = []

    awareness = {
        b: df["Q12"]
        .fillna("")
        .str.lower()
        .str.contains(b.lower(), regex=False)
        .mean()
        * 100
        for b in brands
    }

    top_aw = max(awareness, key=awareness.get)

    insights.append(f"Highest awareness brand: {top_aw}")

    if "Cashify" in awareness:
        if awareness["Cashify"] < max(awareness.values()):
            insights.append(
                "Cashify trails the category leader in awareness"
            )

    insights.append(
        "Influencer and social media channels appear frequently in awareness drivers"
    )

    return insights

def strategic_recommendations(df, brands, nps_cols):

    insights = []

    awareness = {
        b: df["Q12"].fillna("").str.lower().str.contains(
            b.lower(), regex=False
        ).mean()*100
        for b in brands
    }

    leader = max(awareness, key=awareness.get)

    if leader != "Cashify":
        insights.append(
            f"{leader} leads category awareness. Cashify should increase marketing visibility."
        )

    if "Cashify" in brands:
        cash_aw = awareness.get("Cashify",0)

        if cash_aw < 40:
            insights.append(
                "Cashify awareness is below 40%. Stronger digital campaigns are recommended."
            )

    nps_scores = {}

    for b in brands:
        col = nps_cols.get(b,"")
        if col in df.columns:
            nps,_,_,_ = calc_nps(df[col])
            nps_scores[b] = nps

    if nps_scores:
        best_nps = max(nps_scores, key=nps_scores.get)

        if best_nps != "Cashify":
            insights.append(
                f"{best_nps} leads in customer advocacy (NPS). Cashify should improve post-purchase experience."
            )

    if len(insights) == 0:
        insights.append("Cashify performs competitively across major brand metrics.")

    return insights
# ── Main render callback ──────────────────────────────────────────────────────

@app.callback(
    Output("kpi-strip","children"),
    Output("main-chart","figure"),
    Output("section-indicator","children"),

    Input("active-section","data"),
    Input("dd-study","value"),
    Input("dd-gender","value"),
    Input("dd-age","value"),
    Input("dd-city","value"),
    Input("awareness-slider","value")
)
def render(sec, study, gender, age, city, awareness_boost):

    is_r = study == "R"

    df = df_r if is_r else df_b
    brands = BRANDS_R if is_r else BRANDS_B
    lbl = "Refurbished" if is_r else "Buyback"

    g = gender or "All"
    a = age or "All"
    c = city or "All"

    nps_cols = NPS_R if is_r else NPS_B
    fam_cols = FAM_R if is_r else FAM_B
    con_cols = CON_R if is_r else CON_B
    soa_cols = SOA_R if is_r else SOA_B
    cash_drv = CASHIFY_DRV_R if is_r else CASHIFY_DRV_B
    comp_drv = COMP_DRV_R if is_r else COMP_DRV_B

    d = flt(df, g, a, c)
    n = len(d)

    insights = generate_insights(d, brands)
    recommendations = strategic_recommendations(d, brands, nps_cols)
    opportunities = detect_market_opportunities(d, brands)
    channel_insights = channel_strategy_insights(d)

    # KPI calculations
    nps_col = nps_cols.get("Cashify", "")

    if nps_col in d.columns:
        cnps, cpro, cpas, cdet = calc_nps(d[nps_col])
    else:
        cnps, cpro, cpas, cdet = (0, 0, 0, 0)

    c_aw = round(
        d["Q12"]
        .fillna("")
        .str.lower()
        .str.contains("cashify", regex=False)
        .mean()
        * 100
    )

    c_tom = round(
        d["Q10"]
        .fillna("")
        .str.lower()
        .str.contains("cashify", regex=False)
        .mean()
        * 100
    )

    nps_color = "#22C55E" if cnps >= 0 else "#EF4444"

    kpis = dbc.Row(
        [
            kpi_card("Sample", n, CASHIFY, lbl),
            kpi_card("Cashify NPS", cnps, nps_color, "Net Promoter Score"),
            kpi_card("Promoters", f"{cpro}%", "#22C55E", "Score 9–10"),
            kpi_card("Passives", f"{cpas}%", "#F59E0B", "Score 7–8"),
            kpi_card("Detractors", f"{cdet}%", "#EF4444", "Score 0–6"),
            kpi_card("Aided Awareness", f"{c_aw}%", CASHIFY, "Cashify"),
            kpi_card("Top of Mind", f"{c_tom}%", CASHIFY, "Cashify TOM"),
        ]
    )

    # Section indicator
    label_map = {l: name for l, name in SECTION_BUTTONS}

    indicator = [

html.Div([
    html.Span(
        f"{sec}. {label_map.get(sec,'')}",
        style={"color": CASHIFY, "fontWeight": "700", "marginRight": "10px"},
    ),

    html.Span(
        f"{lbl} study | Gender: {g} | Age: {a} | City: {c}",
        style={"color": MUTED},
    ),
]),

html.Div(
    [
        html.Div(
            "Strategic Recommendations",
            style={"fontWeight":"700","marginTop":"8px","fontSize":"13px","color":TEXT},
        ),

        html.Ul(
            [html.Li(r) for r in recommendations],
            style={"color":MUTED,"fontSize":"12px","marginTop":"4px"}
        ),

        html.Div(
            "Key Insights",
            style={"fontWeight":"700","marginTop":"8px","fontSize":"13px","color":TEXT},
        ),

        html.Ul(
            [html.Li(i) for i in insights],
            style={"color":MUTED,"fontSize":"12px"}
        ),
        html.Div(
            "Market Opportunities",
            style={"fontWeight":"700","marginTop":"8px","fontSize":"13px","color":TEXT}
            ),
            
        html.Ul(
            [html.Li(o) for o in opportunities],
            style={"color":MUTED,"fontSize":"12px"}
            ),
        html.Div(
            "Channel Strategy Insights",
            style={"fontWeight":"700","marginTop":"8px","fontSize":"13px","color":TEXT}
            ),
        html.Ul([html.Li(i) for i in (channel_insights or ["No strong channel signals detected"])],
            style={"color":MUTED,"fontSize":"12px"}
        )    
    ]
)
]

    chart_map = {
        "A": lambda: chart_awareness(d, brands, lbl, g, a, c),
        "B": lambda: chart_health(d, brands, fam_cols, lbl, g, a, c),
        "C": lambda: chart_nps(d, brands, nps_cols, lbl, g, a, c),
        "D": lambda: chart_soa(d, brands, soa_cols, lbl, g, a, c),
        "E": lambda: chart_consideration(d, brands, con_cols, lbl, g, a, c),
        "F": lambda: chart_drivers(d, cash_drv, comp_drv, lbl, g, a, c),
        "G": lambda: chart_barriers(d, "Q21B", lbl, g, a, c),
        "H": lambda: chart_category(d, is_r, lbl, g, a, c),
        "I": lambda: chart_brand_funnel(d, brands, fam_cols, con_cols, lbl, g, a, c),
        "J": lambda: chart_positioning(d, brands, nps_cols, fam_cols, con_cols, lbl, g, a, c),
        "K": lambda: chart_radar(d, brands, fam_cols, con_cols, nps_cols, lbl, g, a, c),
        "L": lambda: chart_brand_power(d, brands, fam_cols, con_cols, nps_cols, lbl, g, a, c),
        "M": lambda: chart_market_opportunity(d, brands, lbl, g, a, c),
        "N": lambda: chart_channel_attribution(d, brands, soa_cols, lbl, g, a, c),
        "O": lambda: chart_growth_simulator(d, brands, fam_cols, con_cols, nps_cols, awareness_boost, lbl, g, a, c)
    }

    try:
        fig = chart_map.get(sec, chart_map["A"])()

    except Exception as e:
        print("Chart error:", e)

        fig = go.Figure()

        fig.add_annotation(
            text="Chart could not be generated for this filter.",
            xref="paper",
            yref="paper",
            x=0.5,
            y=0.5,
            showarrow=False,
            font=dict(size=16, color=MUTED),
        )

        fig.update_layout(**LAYOUT, height=400)

    return kpis, fig, indicator

# ── 8. RUN ────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    print("\n" + "="*55)
    print("  Cashify DSS — starting server...")
    print("  Open your browser at: http://127.0.0.1:8050")
    print("="*55 + "\n")
    app.run(debug=False, host="127.0.0.1", port=8050)