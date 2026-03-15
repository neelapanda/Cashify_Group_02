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
    fig.update_layout(**LAYOUT,
                      title=dict(text=title, font=dict(size=14, color=TEXT)),
                      height=h)
    return fig

# ── 4. HELPERS ────────────────────────────────────────────────────────────────
def flt(df, g, a, c):
    d = df.copy()
    if g != "All": d = d[d["Gender"] == g]
    if a != "All": d = d[d["Age"]    == a]
    if c != "All": d = d[d["City"]   == c]
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


def chart_soa(df, brands, soa_cols, lbl, g, a, c):
    d = flt(df, g, a, c); n = len(d)
    all_ch = set()
    for brand in brands:
        col = soa_cols.get(brand, "")
        if col and col in d.columns:
            d[col].dropna().str.split(",").explode().str.strip().apply(
                lambda x: all_ch.add(x) if x and x.lower() not in ("nan","") else None)
    channels = sorted([ch for ch in all_ch if ch and len(ch) > 2])
    matrix, vb = [], []
    for brand in brands:
        col = soa_cols.get(brand, "")
        if not col or col not in d.columns: continue
        row = [round(d[col].fillna("").str.lower()
                     .str.contains(ch.lower(), regex=False).mean()*100, 1)
               for ch in channels]
        matrix.append(row); vb.append(brand)
    if not matrix: return go.Figure()
    fig = go.Figure(go.Heatmap(
        z=matrix, x=channels, y=vb, colorscale="RdYlGn",
        text=[[f"{v}%" for v in row] for row in matrix],
        texttemplate="%{text}", textfont={"size": 9},
        hovertemplate="Brand: %{y}<br>Channel: %{x}<br>%{z}%<extra></extra>"
    ))
    fig.update_layout(**LAYOUT,
                      title=dict(text=f"D.  Source of Awareness Heatmap — {lbl}  (n={n})",
                                 font=dict(size=14, color=TEXT)),
                      height=max(440, 58*len(vb)),
                      xaxis=dict(tickangle=-35, gridcolor=BORDER),
                      yaxis=dict(gridcolor=BORDER))
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
    cs = driver_score(d, cashify_cols)
    ks = driver_score(d, comp_cols)
    all_f = sorted(set(list(cs) + list(ks)))
    if not all_f: return go.Figure()
    mx = max(list(cs.values()) + list(ks.values()) + [1])
    c_n = [round(cs.get(f, 0)/mx*100, 1) for f in all_f]
    k_n = [round(ks.get(f, 0)/mx*100, 1) for f in all_f]
    order  = sorted(range(len(all_f)), key=lambda i: c_n[i])
    fsort  = [all_f[i] for i in order]
    csort  = [c_n[i]   for i in order]
    ksort  = [k_n[i]   for i in order]
    fig = go.Figure()
    fig.add_trace(go.Bar(y=fsort, x=csort, name="Cashify", orientation="h",
                         marker_color=CASHIFY,
                         text=[f"{v}" for v in csort], textposition="outside"))
    fig.add_trace(go.Bar(y=fsort, x=ksort, name="Competitors (avg)", orientation="h",
                         marker_color="#6366F1",
                         text=[f"{v}" for v in ksort], textposition="outside"))
    fig.update_layout(barmode="group")
    return styled(fig, f"F.  Choice Drivers: Cashify vs Competitors — {lbl}  (n={n})", 580)


def chart_barriers(df, col, lbl, g, a, c):
    d = flt(df, g, a, c); n = len(d)
    if col not in d.columns: return go.Figure()
    counts = explode_counts(d, col, 15)
    counts = counts[counts.index.str.len() > 3].sort_values(ascending=True)
    fig = go.Figure(go.Bar(y=counts.index, x=counts.values, orientation="h",
                           marker_color="#EF4444",
                           text=counts.values, textposition="outside"))
    return styled(fig, f"G.  Barriers to Choosing Cashify — {lbl}  (n={n})", 520)


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

# ── 6. DASH LAYOUT ────────────────────────────────────────────────────────────
app = dash.Dash(
    __name__,
    external_stylesheets=[dbc.themes.SLATE],
    title="Cashify DSS",
)

# Shared dropdown style
DD = {"backgroundColor": CARD, "color": TEXT, "border": f"1px solid {BORDER}"}
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
        html.P("Refurbished (Buy) Study  |  Buyback (Sell) Study  |  All 8 Deliverables A–H",
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

    # ── Chart ─────────────────────────────────────────────────────────────
    dcc.Graph(id="main-chart",
              config={"displayModeBar": True, "responsive": True},
              style={"borderRadius": "10px", "overflow": "hidden"}),

    # Hidden store for active section
    dcc.Store(id="active-section", data="A"),
])

# ── 7. CALLBACKS ──────────────────────────────────────────────────────────────

# Update filter options when study changes
@app.callback(
    Output("dd-gender", "options"),
    Output("dd-age",    "options"),
    Output("dd-city",   "options"),
    Output("dd-gender", "value"),
    Output("dd-age",    "value"),
    Output("dd-city",   "value"),
    Input("dd-study", "value"),
)
def update_filter_options(study):
    df = df_r if study == "R" else df_b
    return (dropdown_opts("Gender", df),
            dropdown_opts("Age",    df),
            dropdown_opts("City",   df),
            "All", "All", "All")

# Capture which section button was last clicked → store
@app.callback(
    Output("active-section", "data"),
    [Input(f"btn-{l}", "n_clicks") for l, _ in SECTION_BUTTONS],
    prevent_initial_call=True,
)
def set_section(*args):
    ctx = dash.callback_context
    if not ctx.triggered: return "A"
    btn_id = ctx.triggered[0]["prop_id"].split(".")[0]   # e.g. "btn-B"
    return btn_id.replace("btn-", "")                    # e.g. "B"

# Highlight the active button
@app.callback(
    [Output(f"btn-{l}", "style") for l, _ in SECTION_BUTTONS],
    Input("active-section", "data"),
)
def highlight_button(active):
    styles = []
    for letter, _ in SECTION_BUTTONS:
        if letter == active:
            s = {"background": CASHIFY, "color": "#fff",
                 "border": f"2px solid {CASHIFY}",
                 "borderRadius": "8px", "padding": "8px 16px",
                 "cursor": "pointer", "fontSize": "13px",
                 "fontFamily": "inherit", "fontWeight": "700",
                 "transition": "all 0.15s ease"}
        else:
            s = {"background": CARD, "color": TEXT,
                 "border": f"2px solid {BORDER}",
                 "borderRadius": "8px", "padding": "8px 16px",
                 "cursor": "pointer", "fontSize": "13px",
                 "fontFamily": "inherit", "fontWeight": "500",
                 "transition": "all 0.15s ease"}
        styles.append(s)
    return styles

# Main render callback: KPIs + chart + section label
@app.callback(
    Output("kpi-strip",          "children"),
    Output("main-chart",         "figure"),
    Output("section-indicator",  "children"),
    Input("active-section", "data"),
    Input("dd-study",  "value"),
    Input("dd-gender", "value"),
    Input("dd-age",    "value"),
    Input("dd-city",   "value"),
)
def render(sec, study, gender, age, city):
    is_r   = (study == "R")
    df     = df_r if is_r else df_b
    brands = BRANDS_R if is_r else BRANDS_B
    lbl    = "Refurbished" if is_r else "Buyback"
    g, a, c = gender or "All", age or "All", city or "All"

    nps_cols  = NPS_R if is_r else NPS_B
    fam_cols  = FAM_R if is_r else FAM_B
    con_cols  = CON_R if is_r else CON_B
    soa_cols  = SOA_R if is_r else SOA_B
    cash_drv  = CASHIFY_DRV_R if is_r else CASHIFY_DRV_B
    comp_drv  = COMP_DRV_R    if is_r else COMP_DRV_B

    d = flt(df, g, a, c)
    n = len(d)

    # ── KPI calculations ─────────────────────────────────────────────────────
    nps_col = nps_cols.get("Cashify", "")
    cnps, cpro, cpas, cdet = (calc_nps(d[nps_col])
                               if nps_col in d.columns else (0,0,0,0))
    c_aw  = round(d["Q12"].fillna("").str.lower()
                  .str.contains("cashify", regex=False).mean()*100)
    c_tom = round(d["Q10"].fillna("").str.lower()
                  .str.contains("cashify", regex=False).mean()*100)

    nps_color = "#22C55E" if cnps >= 0 else "#EF4444"
    kpis = dbc.Row([
        kpi_card("Sample",         n,            CASHIFY,      lbl),
        kpi_card("Cashify NPS",    cnps,         nps_color,    "Net Promoter Score"),
        kpi_card("Promoters",      f"{cpro}%",   "#22C55E",    "Score 9–10"),
        kpi_card("Passives",       f"{cpas}%",   "#F59E0B",    "Score 7–8"),
        kpi_card("Detractors",     f"{cdet}%",   "#EF4444",    "Score 0–6"),
        kpi_card("Aided Awareness",f"{c_aw}%",   CASHIFY,      "Cashify (Q12)"),
        kpi_card("Top of Mind",    f"{c_tom}%",  CASHIFY,      "Cashify TOM"),
    ])

    # ── Section banner ────────────────────────────────────────────────────────
    label_map = {l: name for l, name in SECTION_BUTTONS}
    indicator = [
        html.Span(f"{sec}. {label_map.get(sec,'')}",
                  style={"color": CASHIFY, "fontWeight": "700",
                         "marginRight": "10px"}),
        html.Span(f"{lbl} study  |  Gender: {g}  |  Age: {a}  |  City: {c}",
                  style={"color": MUTED}),
    ]

    # ── Chart ─────────────────────────────────────────────────────────────────
    chart_map = {
        "A": lambda: chart_awareness(df, brands, lbl, g, a, c),
        "B": lambda: chart_health(df, brands, fam_cols, lbl, g, a, c),
        "C": lambda: chart_nps(df, brands, nps_cols, lbl, g, a, c),
        "D": lambda: chart_soa(df, brands, soa_cols, lbl, g, a, c),
        "E": lambda: chart_consideration(df, brands, con_cols, lbl, g, a, c),
        "F": lambda: chart_drivers(df, cash_drv, comp_drv, lbl, g, a, c),
        "G": lambda: chart_barriers(df, "Q21B", lbl, g, a, c),
        "H": lambda: chart_category(df, is_r, lbl, g, a, c),
    }
    fig = chart_map.get(sec, chart_map["A"])()

    return kpis, fig, indicator


# ── 8. RUN ────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    print("\n" + "="*55)
    print("  Cashify DSS — starting server...")
    print("  Open your browser at: http://127.0.0.1:8050")
    print("="*55 + "\n")
    app.run(debug=False, host="127.0.0.1", port=8050)
