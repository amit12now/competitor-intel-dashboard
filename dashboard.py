#!/usr/bin/env python3
"""
dashboard.py
------------
Competitor Intelligence Dashboard (Air Compressors) - Atlas Copco

Reads ONLY the Clean_* sheets (+ Settings) of Competitor_Intelligence_Master.xlsx,
so it always reflects whatever the ETL pipeline last produced. Run with:

    streamlit run dashboard.py

By default it looks for Competitor_Intelligence_Master.xlsx in the same folder as
this script. You can also upload a different copy from the sidebar.

IMPORTANT - what this dashboard is, and isn't:
This dashboard surfaces external competitor SIGNALS for review. It does not make
final business decisions. Every section uses soft, review-oriented language
("may be worth reviewing", "could indicate", "possible opportunity") rather than
directives. Categories with zero confirmed signals are shown honestly as
"No signals found this period" rather than estimated or invented.

Layout: one Executive Summary tab, one profile tab per top-activity competitor,
one summary tab for the remaining competitors, and a Raw Data Explorer with
filters. Ingersoll Rand is always pinned as the first competitor tab; the other
top-tab slots are ranked live from the data each run. Logos and LinkedIn post
images are loaded from small local lookup files (competitor_logos.json,
post_images.json) built from the scraped source data - if a logo or image isn't
available for a given competitor, a plain letter badge / text-only card is shown
instead rather than inventing one.
"""
from pathlib import Path
import json

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ---------------------------------------------------------------------------
# Brand palette (Atlas Copco)
# ---------------------------------------------------------------------------
ATLAS_BLUE = "#0099CC"
ATLAS_BLUE_DARK = "#006C92"
ATLAS_BLUE_LIGHT = "#E6F5FA"
NAVY = "#1B1F2A"
GRAY = "#6B7280"
GRAY_LIGHT = "#F4F6F8"
WHITE = "#FFFFFF"
ACCENT_AMBER = "#E8A33D"
ACCENT_AMBER_LIGHT = "#FFF6E8"
ACCENT_RED = "#D8554A"
ACCENT_GREEN = "#3FA66A"
NEUTRAL_BAR = "#C7CDD4"

CATEGORICAL_PALETTE = [
    ATLAS_BLUE, ACCENT_AMBER, ACCENT_GREEN, ACCENT_RED, "#7C5CBF",
    "#3D8C97", "#C97FB0", "#5B7FA6", "#A3A86C", "#8D6E63", "#4F6D7A",
]
PRIORITY_COLORS = {"High": ACCENT_RED, "Medium": ACCENT_AMBER, "Low": "#9AA5B1"}
PRIORITY_ICONS = {"High": "\U0001F534", "Medium": "\U0001F7E1", "Low": "\U0001F7E2"}
CONFIDENCE_COLORS = {"High": ACCENT_GREEN, "Medium": ACCENT_AMBER, "Low": "#9AA5B1"}
SENTIMENT_COLORS = {"Positive": ACCENT_GREEN, "Neutral": "#9AA5B1", "Mixed": ACCENT_AMBER, "Negative": ACCENT_RED}
ACTIVITY_LEVEL_COLORS = {"High": ATLAS_BLUE, "Medium": ACCENT_AMBER, "Low": NEUTRAL_BAR}

DISCLAIMER = (
    "This dashboard surfaces external competitor signals for review. Final actions "
    "should be validated by the relevant marketing, product, sales, or leadership teams."
)

DEFAULT_PATH = Path(__file__).parent / "Competitor_Intelligence_Master.xlsx"
LOGO_ASSET_PATH = Path(__file__).parent / "competitor_logos.json"
IMAGE_ASSET_PATH = Path(__file__).parent / "post_images.json"

CLEAN_SHEETS = [
    "Clean_Executive_Summary", "Clean_Competitor_Activity", "Clean_Channel_Activity",
    "Clean_Competitor_x_Channel", "Clean_Trend_May_vs_June", "Clean_Theme_Frequency",
    "Clean_Theme_by_Competitor", "Clean_Theme_by_Channel", "Clean_Theme_Trend_May_vs_June",
    "Clean_Product_Service_Signals", "Clean_Social_Conversation", "Clean_Conversation_Top_Themes",
    "Clean_Conversation_Sentiment", "Clean_Most_Mentioned_Brands", "Clean_PR_News_Events",
    "Clean_Hiring_Expansion", "Clean_Opportunities", "Clean_Raw_Data",
]

CHANNEL_LABELS = {
    "LinkedIn posts": "LinkedIn", "Instagram posts": "Instagram", "YouTube videos": "YouTube",
    "Blogs": "Blogs", "Events": "Events", "Webinars": "Webinars", "PR releases": "PR releases",
    "News mentions": "News mentions", "LinkedIn jobs": "LinkedIn jobs",
    "Employee/person posts": "Employee posts", "New product/service pages": "Product pages",
}

st.set_page_config(page_title="Competitor Intelligence Dashboard | Atlas Copco", layout="wide",
                    initial_sidebar_state="expanded")

# ---------------------------------------------------------------------------
# Brand styling
# ---------------------------------------------------------------------------
st.markdown(f"""
<style>
.stApp {{ background-color: {GRAY_LIGHT}; }}
section[data-testid="stSidebar"] {{ background-color: {NAVY}; }}
section[data-testid="stSidebar"] * {{ color: {WHITE} !important; }}
h1, h2, h3 {{ color: {NAVY}; font-family: "Segoe UI", Calibri, sans-serif; }}
div[data-testid="stMetric"] {{
    background-color: {WHITE};
    border: 1px solid #E2E8F0;
    border-left: 5px solid {ATLAS_BLUE};
    border-radius: 8px;
    padding: 12px 14px 8px 14px;
    box-shadow: 0 1px 3px rgba(0,0,0,0.04);
}}
div[data-testid="stMetricLabel"] {{ color: {GRAY}; font-size: 0.80rem; }}
div[data-testid="stMetricValue"] {{ color: {NAVY}; }}
.theme-pill {{
    display:inline-block; background:{ATLAS_BLUE_LIGHT}; color:{ATLAS_BLUE_DARK};
    border-radius: 14px; padding: 3px 12px; margin: 2px; font-size:0.82rem; font-weight:600;
}}
.stTabs [data-baseweb="tab-list"] {{
    gap: 6px; background-color: {WHITE}; padding: 6px; border-radius: 12px;
    border: 1px solid #E2E8F0; flex-wrap: wrap;
}}
.stTabs [data-baseweb="tab"] {{
    background-color: {GRAY_LIGHT}; border-radius: 9px; padding: 10px 18px;
    color: {NAVY}; font-weight: 600; border: none;
}}
.stTabs [data-baseweb="tab"]:hover {{ background-color: {ATLAS_BLUE_LIGHT}; color: {ATLAS_BLUE_DARK}; }}
.stTabs [aria-selected="true"] {{
    background-color: {ATLAS_BLUE} !important; color: {WHITE} !important; font-weight: 700;
}}
.stTabs [data-baseweb="tab-highlight"] {{ background-color: transparent !important; }}
.stTabs [data-baseweb="tab-border"] {{ display: none; }}
.main-header {{
    display: flex; align-items: center; gap: 16px;
    background: linear-gradient(135deg, {ATLAS_BLUE} 0%, {ATLAS_BLUE_DARK} 100%);
    color: {WHITE};
    padding: 22px 28px;
    border-radius: 14px;
    margin-bottom: 14px;
}}
.main-header-icon {{ font-size: 2.1rem; line-height: 1; }}
.main-header-title {{ font-size: 1.55rem; font-weight: 800; color: {WHITE}; line-height: 1.25; }}
.main-header-sub {{ font-size: 0.92rem; color: rgba(255,255,255,0.88); margin-top: 3px; }}
.kpi-card {{
    background-color: {WHITE};
    border: 1px solid #E2E8F0;
    border-left: 4px solid {ATLAS_BLUE};
    border-radius: 10px;
    padding: 14px 16px 12px 16px;
    height: 100%;
    box-shadow: 0 1px 3px rgba(0,0,0,0.04);
}}
.kpi-icon {{ font-size: 1.25rem; margin-bottom: 6px; }}
.kpi-label {{
    color: {GRAY}; font-size: 0.72rem; font-weight: 700;
    text-transform: uppercase; letter-spacing: 0.04em;
}}
.kpi-value {{ color: {NAVY}; font-size: 1.55rem; font-weight: 800; margin-top: 2px; line-height: 1.2; }}
.kpi-sub {{ font-size: 0.76rem; margin-top: 4px; font-weight: 600; color: {GRAY}; }}
.story-banner {{
    background: linear-gradient(135deg, {ATLAS_BLUE} 0%, {ATLAS_BLUE_DARK} 100%);
    color: {WHITE};
    padding: 22px 26px;
    border-radius: 12px;
    font-size: 1.08rem;
    line-height: 1.55;
    margin-bottom: 6px;
}}
.story-banner b {{ color: {WHITE}; }}
.elaborate-box {{
    background-color: {WHITE};
    border: 1px solid #E2E8F0;
    border-radius: 10px;
    padding: 16px 20px;
    font-size: 0.95rem;
    line-height: 1.6;
    color: {NAVY};
    margin-bottom: 16px;
}}
.elaborate-box b {{ color: {ATLAS_BLUE_DARK}; }}
.insight-box {{
    background-color: {ACCENT_AMBER_LIGHT};
    border-left: 5px solid {ACCENT_AMBER};
    border-radius: 6px;
    padding: 10px 16px;
    margin: 4px 0 16px 0;
    font-size: 0.90rem;
    color: {NAVY};
    line-height: 1.5;
}}
.insight-box b {{ color: {ATLAS_BLUE_DARK}; }}
.action-box {{
    background-color: {ATLAS_BLUE_LIGHT};
    border-left: 5px solid {ATLAS_BLUE};
    border-radius: 6px;
    padding: 10px 16px;
    margin: 0 0 22px 0;
    font-size: 0.90rem;
    color: {NAVY};
    line-height: 1.5;
}}
.action-box b {{ color: {ATLAS_BLUE_DARK}; }}
.mini-stat {{
    background-color: {WHITE};
    border: 1px solid #E2E8F0;
    border-radius: 10px;
    padding: 14px 16px;
    height: 100%;
}}
.mini-stat .label {{ color: {GRAY}; font-size: 0.78rem; font-weight: 600; text-transform: uppercase; letter-spacing: 0.03em; }}
.mini-stat .value {{ color: {NAVY}; font-size: 1.05rem; font-weight: 700; margin-top: 4px; line-height: 1.4; }}
.next-step {{
    background-color: {WHITE};
    border: 1px solid #E2E8F0;
    border-left: 5px solid {ATLAS_BLUE};
    border-radius: 8px;
    padding: 10px 16px;
    margin-bottom: 8px;
    font-size: 0.92rem;
    color: {NAVY};
}}
.chip {{
    display:inline-block; border-radius: 10px; padding: 2px 10px; margin: 0 4px 6px 0;
    font-size: 0.74rem; font-weight: 700;
}}
.post-meaning {{ color: {GRAY}; font-size: 0.82rem; }}
.section-hint {{ color: {GRAY}; font-size: 0.78rem; font-weight: 400; margin-left: 4px; }}
.takeaway-box {{
    background-color: {WHITE};
    border: 1px solid #E2E8F0;
    border-radius: 10px;
    padding: 4px 18px;
    margin-bottom: 18px;
}}
.takeaway-item {{
    display: flex; gap: 10px; align-items: flex-start;
    padding: 11px 0; border-bottom: 1px solid {GRAY_LIGHT};
    font-size: 0.88rem; color: {NAVY}; line-height: 1.5;
}}
.takeaway-item:last-child {{ border-bottom: none; }}
.takeaway-ic {{ flex: 0 0 auto; font-size: 1rem; line-height: 1.4; }}
.takeaway-item b {{ color: {ATLAS_BLUE_DARK}; }}
.kpi-card, .mini-stat, .next-step {{ transition: box-shadow 0.15s ease; }}
.kpi-card:hover, .mini-stat:hover {{ box-shadow: 0 4px 12px rgba(0,0,0,0.08); }}
</style>
""", unsafe_allow_html=True)


def style_layout(fig, height=420, legend="bottom"):
    """Apply consistent brand styling to every chart.

    legend: "bottom"  -> compact horizontal legend below the plot (default; never
                          competes with the title for space at the top)
            "right"   -> vertical legend to the right (used for donuts, where the
                          reference style is a clean side legend, not labels jammed
                          onto the slices)
            "none"    -> no legend at all (used when color is purely a "leader vs.
                          the rest" highlight, or when the legend would just repeat
                          text already on the axis)
    """
    if legend == "right":
        legend_cfg = dict(orientation="v", yanchor="top", y=1, xanchor="left", x=1.02, font=dict(size=10))
        margin = dict(l=10, r=150, t=48, b=10)
    elif legend == "none":
        legend_cfg = dict()
        margin = dict(l=10, r=28, t=48, b=10)
    else:
        legend_cfg = dict(orientation="h", yanchor="top", y=-0.24, xanchor="left", x=0, font=dict(size=11))
        margin = dict(l=10, r=24, t=48, b=64)
    fig.update_layout(
        plot_bgcolor=WHITE, paper_bgcolor=WHITE,
        font=dict(family="Segoe UI, Calibri, sans-serif", color=NAVY, size=12),
        legend=legend_cfg,
        showlegend=(legend != "none"),
        margin=margin,
        title_font=dict(size=15, color=NAVY),
        height=height,
        uniformtext=dict(mode="hide", minsize=9),
        bargap=0.30,
    )
    return fig


def clean_value_axis(fig, orientation="h"):
    """Hide the numeric value axis when bars already carry an outside text label
    (so the number only appears once, not on the bar AND the axis), and strip
    gridlines everywhere for a flatter, calmer look."""
    if orientation == "h":
        fig.update_xaxes(visible=False, showgrid=False, zeroline=False)
        fig.update_yaxes(showgrid=False, zeroline=False, showline=False, ticks="")
    else:
        fig.update_yaxes(visible=False, showgrid=False, zeroline=False)
        fig.update_xaxes(showgrid=False, zeroline=False, showline=False, ticks="")
    return fig


def highlight_leader(df, value_col, flag_col="_Highlight"):
    """Mark the top row of a single-series bar chart as 'Leading' so it can be
    colored differently from the rest - the same "one bar stands out, the rest
    are neutral gray" treatment used throughout, instead of a same-color bar
    chart with a legend nobody needs."""
    d = df.copy()
    top = d[value_col].max()
    d[flag_col] = [("Leading" if v == top else "Other") for v in d[value_col]]
    return d


# ---------------------------------------------------------------------------
# Small text / UI helpers (for plain-English, data-grounded storytelling)
# ---------------------------------------------------------------------------
def _join_natural(items):
    items = [str(i) for i in items if str(i).strip()]
    if not items:
        return ""
    if len(items) == 1:
        return items[0]
    if len(items) == 2:
        return f"{items[0]} and {items[1]}"
    return ", ".join(items[:-1]) + f", and {items[-1]}"


def _pct(n, d):
    return f"{(n / d * 100):.0f}%" if d else "0%"


def _safe(value, default=""):
    """Coerce a possibly-missing/NaN cell into clean display text, never the
    literal string 'nan' or 'None'."""
    if value is None:
        return default
    try:
        if pd.isna(value):
            return default
    except (TypeError, ValueError):
        pass
    text = str(value).strip()
    if text.lower() in ("nan", "none", ""):
        return default
    return text


def insight(text):
    st.markdown(f"<div class='insight-box'>\U0001F4A1 <b>What this means:</b> {text}</div>", unsafe_allow_html=True)


def action(text):
    st.markdown(f"<div class='action-box'>✅ <b>Worth considering:</b> {text}</div>", unsafe_allow_html=True)


def mini_stat(col, label, value):
    col.markdown(
        f"<div class='mini-stat'><div class='label'>{label}</div><div class='value'>{value}</div></div>",
        unsafe_allow_html=True,
    )


def kpi_card(col, icon, label, value, sublabel=None):
    sub_html = f"<div class='kpi-sub'>{sublabel}</div>" if sublabel else ""
    col.markdown(
        f"<div class='kpi-card'><div class='kpi-icon'>{icon}</div>"
        f"<div class='kpi-label'>{label}</div><div class='kpi-value'>{value}</div>{sub_html}</div>",
        unsafe_allow_html=True,
    )


def chip(text, bg, fg=WHITE):
    return f"<span class='chip' style='background:{bg};color:{fg};'>{text}</span>"


def logo_badge(name, logo_map, size=64):
    """Render a competitor's logo if we have a scraped URL for it, otherwise a
    plain colored letter-badge - never a fabricated/placeholder logo image."""
    url = (logo_map or {}).get(name)
    if url:
        return (f"<img src='{url}' style='width:{size}px;height:{size}px;border-radius:50%;"
                f"object-fit:cover;border:2px solid #E2E8F0;box-shadow:0 1px 3px rgba(0,0,0,0.08);' />")
    initials = "".join(w[0] for w in name.split()[:2]).upper() or "?"
    color = CATEGORICAL_PALETTE[sum(ord(c) for c in name) % len(CATEGORICAL_PALETTE)]
    fsize = max(12, int(size * 0.34))
    return (f"<div style='width:{size}px;height:{size}px;border-radius:50%;background:{color};"
            f"display:flex;align-items:center;justify-content:center;color:white;font-weight:800;"
            f"font-size:{fsize}px;'>{initials}</div>")


# ---------------------------------------------------------------------------
# Data loading
# ---------------------------------------------------------------------------
@st.cache_data(show_spinner="Loading competitor intelligence data...")
def load_data(file_bytes, mtime):
    xls = pd.ExcelFile(file_bytes)
    sheets = {}
    for name in CLEAN_SHEETS:
        try:
            sheets[name] = pd.read_excel(xls, sheet_name=name, header=2)
        except Exception:
            sheets[name] = pd.DataFrame()
    try:
        settings = pd.read_excel(xls, sheet_name="Settings", header=2)
        settings = dict(zip(settings["Setting"], settings["Value"]))
    except Exception:
        settings = {}
    return sheets, settings


@st.cache_data(show_spinner=False)
def load_json_asset(path_str):
    p = Path(path_str)
    if not p.exists():
        return {}
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return {}


def info_if_empty(df, label):
    if df is None or df.empty:
        st.info(f"No signals found this period for {label}.")
        return True
    return False


def priority_badge_table(df, cols=None):
    """Render a dataframe with Priority/Confidence/Sentiment columns color-styled."""
    d = df.copy() if cols is None else df[cols].copy()

    def _color(col, palette):
        return [f"background-color: {palette.get(v, WHITE)}; color: white; font-weight:600;"
                if v in palette else "" for v in d[col]]

    styler = d.style
    if "Priority" in d.columns:
        styler = styler.apply(lambda s: _color("Priority", PRIORITY_COLORS), subset=["Priority"])
    if "Confidence" in d.columns:
        styler = styler.apply(lambda s: _color("Confidence", CONFIDENCE_COLORS), subset=["Confidence"])
    if "Sentiment" in d.columns:
        styler = styler.apply(lambda s: _color("Sentiment", SENTIMENT_COLORS), subset=["Sentiment"])
    return styler


# ---------------------------------------------------------------------------
# Sidebar
# ---------------------------------------------------------------------------
st.sidebar.markdown("## \U0001F9ED Competitor Intelligence")
st.sidebar.caption("Air Compressor Category | Atlas Copco")
upload = st.sidebar.file_uploader("Upload a different workbook copy", type=["xlsx"])
source = upload if upload is not None else (str(DEFAULT_PATH) if DEFAULT_PATH.exists() else None)

if source is None:
    st.warning(
        f"Couldn't find {DEFAULT_PATH.name} next to dashboard.py. "
        "Upload a copy from the sidebar, or place the workbook in this folder."
    )
    st.stop()

mtime = upload.size if upload is not None else DEFAULT_PATH.stat().st_mtime
sheets, settings = load_data(source, mtime)
LOGO_MAP = load_json_asset(str(LOGO_ASSET_PATH))
IMAGE_MAP = load_json_asset(str(IMAGE_ASSET_PATH))

st.sidebar.markdown("---")
st.sidebar.markdown(f"**Report month:** {settings.get('Report_Month', 'n/a')}")
st.sidebar.markdown(f"**Baseline month:** {settings.get('Baseline_Month', 'n/a')}")
st.sidebar.markdown(f"**Our brand:** {settings.get('Our_Brand', 'n/a')}")
st.sidebar.markdown(f"**Competitors tracked:** {len(str(settings.get('Competitors_Tracked','')).split(';'))}")
st.sidebar.markdown(f"**Generated:** {settings.get('Generated_Date', 'n/a')}")
with st.sidebar.expander("Methodology notes"):
    for k in ["Methodology_Note_1", "Methodology_Note_2", "Methodology_Note_3"]:
        if settings.get(k):
            st.caption(settings.get(k))

exec_df = sheets["Clean_Executive_Summary"]
overview = sheets["Clean_Competitor_Activity"]
channel_df = sheets["Clean_Channel_Activity"]
comp_x_chan = sheets["Clean_Competitor_x_Channel"]
trend_df = sheets["Clean_Trend_May_vs_June"]
theme_freq = sheets["Clean_Theme_Frequency"]
theme_by_comp = sheets["Clean_Theme_by_Competitor"]
theme_by_chan = sheets["Clean_Theme_by_Channel"]
theme_trend = sheets["Clean_Theme_Trend_May_vs_June"]
product_signals = sheets["Clean_Product_Service_Signals"]
social_conv = sheets["Clean_Social_Conversation"]
conv_top_themes = sheets["Clean_Conversation_Top_Themes"]
conv_sentiment = sheets["Clean_Conversation_Sentiment"]
brand_mentions = sheets["Clean_Most_Mentioned_Brands"]
pr_news = sheets["Clean_PR_News_Events"]
hiring = sheets["Clean_Hiring_Expansion"]
opportunities = sheets["Clean_Opportunities"]
raw_data = sheets["Clean_Raw_Data"]


def exec_metric(label):
    row = exec_df[exec_df["Metric"] == label]
    return row["Value"].iloc[0] if not row.empty else "n/a"


# ---------------------------------------------------------------------------
# Rank competitors live from the data - top 4 by activity get their own tab,
# the rest are summarized in one comparison table. The one deliberate
# exception: Ingersoll Rand is always pinned as tab #1 per explicit request,
# regardless of its activity ranking that period. The remaining 3 top slots
# are still chosen live by activity among everyone else.
# ---------------------------------------------------------------------------
PINNED_FIRST_COMPETITOR = "Ingersoll Rand"

if not overview.empty:
    _ranked = overview.sort_values("Total activity", ascending=False).reset_index(drop=True)
    if PINNED_FIRST_COMPETITOR in _ranked["Competitor"].values:
        _rest_ranked = _ranked.loc[_ranked["Competitor"] != PINNED_FIRST_COMPETITOR, "Competitor"]
        TOP_COMPETITORS = [PINNED_FIRST_COMPETITOR] + _rest_ranked.head(3).tolist()
        OTHER_COMPETITORS = _rest_ranked.iloc[3:].tolist()
    else:
        TOP_COMPETITORS = _ranked["Competitor"].head(4).tolist()
        OTHER_COMPETITORS = _ranked["Competitor"].iloc[4:].tolist()
else:
    TOP_COMPETITORS, OTHER_COMPETITORS = [], []


def _competitor_row(name):
    r = overview.loc[overview["Competitor"] == name]
    return r.iloc[0] if not r.empty else None


def _top_channel_for(row):
    chan_cols = [c for c in CHANNEL_LABELS if c in overview.columns and c not in
                 ("New product/service pages",)]
    counts = {c: row[c] for c in chan_cols if pd.notna(row.get(c))}
    if not counts or max(counts.values()) == 0:
        return "—"
    best = max(counts, key=counts.get)
    return CHANNEL_LABELS.get(best, best)


def _top_theme_for(name):
    if theme_by_comp.empty or name not in theme_by_comp["Competitor"].values:
        return "—"
    trow = theme_by_comp.loc[theme_by_comp["Competitor"] == name].iloc[0]
    tcols = [c for c in theme_by_comp.columns if c != "Competitor"]
    tvals = trow[tcols]
    if tvals.empty or tvals.max() == 0:
        return "—"
    return tvals.idxmax()



def render_post_card(p):
    """Compact, image-forward card for the post grid. Channel is implied by
    the section it's grouped under, so we don't repeat it on the card itself -
    keeps each card scannable at a glance instead of a paragraph of text."""
    with st.container(border=True):
        img_urls = IMAGE_MAP.get(_safe(p.get("URL")), [])
        if img_urls:
            try:
                st.image(img_urls[0], width="stretch")
            except Exception:
                pass

        sent = _safe(p.get("Sentiment"))
        theme = _safe(p.get("Theme"))
        chips_html = ""
        if theme:
            chips_html += chip(theme, "#6B7280")
        if sent:
            chips_html += chip(sent, SENTIMENT_COLORS.get(sent, GRAY))
        if chips_html:
            st.markdown(chips_html, unsafe_allow_html=True)

        title = _safe(p.get("Title")) or _safe(p.get("Summary"))[:80]
        if title:
            st.markdown(f"**{title[:90]}**")

        dt = pd.to_datetime(p.get("Date"), errors="coerce")
        date_str = dt.strftime("%b %d, %Y") if pd.notna(dt) else ""
        eng = p.get("Engagement")
        eng_str = f"{int(eng):,} engagement" if pd.notna(eng) else ""
        meta_bits = [b for b in [date_str, eng_str] if b]
        if meta_bits:
            st.caption(" · ".join(meta_bits))
        url = _safe(p.get("URL"))
        if url:
            st.markdown(f"[View source ↗]({url})")


# Content the competitor itself published, broken out by channel - shown as
# separate subsections rather than merged into one feed (explicit requirement).
# Third tuple element is the icon used in the coverage strip and section headers.
OWNED_CHANNEL_BUCKETS = [
    ("LinkedIn", {"LinkedIn company posts", "LinkedIn posts from company employees"}, "\U0001F4BC"),
    ("Instagram", {"Instagram posts"}, "\U0001F4F7"),
    ("YouTube", {"YouTube videos"}, "▶️"),
    ("Blog", {"Blog pages/articles"}, "\U0001F4DD"),
]
OWNED_SOURCE_TYPES = {t for _, types, _ in OWNED_CHANNEL_BUCKETS for t in types}

# Other people's posts/comments that merely mention the competitor's name,
# also broken out by type rather than merged into one feed.
MENTION_TYPE_BUCKETS = [
    ("Posts mentioning their name", {"LinkedIn posts mentioning competitor names"}, "\U0001F5E3️"),
    ("Market-reaction comments", {"LinkedIn comments (market reaction)"}, "\U0001F4AC"),
    ("Posts mentioning our brand", {"LinkedIn posts mentioning our brand"}, "\U0001F3F7️"),
    ("Posts mentioning ‘air compressor’", {"LinkedIn posts mentioning 'air compressor'"}, "\U0001F527"),
]
MENTION_SOURCE_TYPES = {t for _, types, _ in MENTION_TYPE_BUCKETS for t in types}


def render_post_section(posts_df, limit=4, cols=2):
    """Render up to `limit` most recent posts from posts_df as a card grid
    (side by side, `cols` per row) instead of one full-width card per row -
    keeps a handful of posts looking like a gallery, not a scrolling feed.
    Callers should skip this entirely when posts_df is empty."""
    if posts_df.empty:
        return
    posts_df = posts_df.sort_values("_dt", ascending=False).head(limit)
    rows = list(posts_df.iterrows())
    for i in range(0, len(rows), cols):
        row_slice = rows[i:i + cols]
        columns = st.columns(cols)
        for col, (_, p) in zip(columns, row_slice):
            with col:
                render_post_card(p)


def _theme_color_map(theme_series):
    """Stable theme -> brand-palette color mapping, ordered by how often each
    theme appears (most-frequent theme gets the first palette color), so the
    same theme keeps the same color across the engagement-over-time, by-theme,
    and posting-mix charts below - a visual thread tying the three together."""
    order = theme_series.value_counts().index.tolist()
    return {t: CATEGORICAL_PALETTE[i % len(CATEGORICAL_PALETTE)] for i, t in enumerate(order)}


def render_engagement_section(name, owned_posts):
    """Premium 'content performance' block for a competitor's OWN posts only
    (not third-party mentions) - engagement over time, engagement by theme,
    posting mix by theme, and a sentiment mix. Built entirely from the real
    Engagement/Theme/Sentiment/Date columns already in Clean_Raw_Data. We do
    NOT fabricate a reaction-type breakdown (Like/Praise/Empathy counts) since
    the scraped data doesn't contain one - sentiment classification is the
    closest real signal we have, so that's what powers the fourth chart."""
    if owned_posts.empty:
        return
    d = owned_posts.copy()
    d["Engagement"] = pd.to_numeric(d["Engagement"], errors="coerce")
    d["_dt"] = pd.to_datetime(d["Date"], errors="coerce")
    d = d.dropna(subset=["Engagement", "_dt"]).sort_values("_dt")
    if d.empty:
        return

    theme_colors = _theme_color_map(d["Theme"])

    st.markdown(
        "##### \U0001F4C8 Engagement & content performance"
        f"<span class='section-hint'>· {name}'s own posts only, not third-party mentions</span>",
        unsafe_allow_html=True,
    )

    total_posts = len(d)
    avg_eng = d["Engagement"].mean()
    days_span = max((d["_dt"].max() - d["_dt"].min()).days, 1)
    cadence = total_posts / (days_span / 7)
    best = d.loc[d["Engagement"].idxmax()]

    e1, e2, e3, e4 = st.columns(4)
    kpi_card(e1, "\U0001F4C5", "Posts in window", total_posts, f"{cadence:.1f} posts / week")
    kpi_card(e2, "\U0001F525", "Avg engagement / post", f"{avg_eng:,.0f}")
    kpi_card(e3, "\U0001F3C6", "Top post engagement", f"{int(best['Engagement']):,}",
             best["_dt"].strftime("%b %d, %Y"))
    kpi_card(e4, "\U0001F4AC", "Total engagement", f"{int(d['Engagement'].sum()):,}")
    st.write("")

    g1, g2 = st.columns(2)
    with g1:
        fig = go.Figure(go.Scatter(
            x=d["_dt"], y=d["Engagement"], mode="lines+markers",
            line=dict(color=ATLAS_BLUE, width=2), fill="tozeroy", fillcolor="rgba(0,153,204,0.12)",
            marker=dict(size=9, color=[theme_colors.get(t, GRAY) for t in d["Theme"]],
                        line=dict(width=1, color=WHITE)),
            customdata=d["Theme"].values,
            hovertemplate="%{x|%b %d, %Y}<br>Engagement: %{y}<br>Theme: %{customdata}<extra></extra>",
        ))
        fig.update_layout(title="Engagement over time")
        fig.update_xaxes(showgrid=False)
        fig.update_yaxes(showgrid=True, gridcolor="#EEF1F4", zeroline=False, rangemode="tozero")
        st.plotly_chart(style_layout(fig, 340, legend="none"), width="stretch")
        st.caption("\U0001F4A1 Dot color = content theme; hover a point for that post's date and theme.")
    with g2:
        theme_avg = d.groupby("Theme")["Engagement"].mean().sort_values(ascending=True)
        tdf = pd.DataFrame({"Theme": theme_avg.index, "Avg engagement": theme_avg.values})
        fig = px.bar(tdf, x="Avg engagement", y="Theme", orientation="h", text="Avg engagement",
                     color="Theme", color_discrete_map=theme_colors, title="Engagement by content theme")
        fig.update_yaxes(title="")
        fig.update_traces(textposition="outside", cliponaxis=False, texttemplate="%{text:.0f}")
        clean_value_axis(fig, "h")
        st.plotly_chart(style_layout(fig, 340, legend="none"), width="stretch")
        st.caption("\U0001F4A1 Average engagement per post, by theme - what resonates, not just what's posted most.")

    g3, g4 = st.columns(2)
    with g3:
        mix = d["Theme"].value_counts()
        fig = px.pie(names=mix.index, values=mix.values, hole=0.5,
                     color=mix.index, color_discrete_map=theme_colors, title="Posting mix by theme")
        fig.update_traces(textinfo="percent", textfont_size=11)
        st.plotly_chart(style_layout(fig, 340, legend="right"), width="stretch")
        st.caption("\U0001F4A1 Share of their own posts by theme - what they choose to post about.")
    with g4:
        sent_counts = d["Sentiment"].value_counts()
        sent_order = [s for s in ["Positive", "Neutral", "Mixed", "Negative"] if s in sent_counts.index]
        sdf = pd.DataFrame({"Sentiment": sent_order, "Posts": [sent_counts[s] for s in sent_order]})
        fig = px.bar(sdf, x="Sentiment", y="Posts", text="Posts", color="Sentiment",
                     color_discrete_map=SENTIMENT_COLORS, title="Sentiment mix")
        fig.update_traces(textposition="outside", cliponaxis=False)
        fig.update_xaxes(title="")
        clean_value_axis(fig, "v")
        st.plotly_chart(style_layout(fig, 340, legend="none"), width="stretch")
        st.caption("\U0001F4A1 Sentiment we classified on their own posts (LinkedIn doesn't expose a public "
                   "reaction-type breakdown, so we use our sentiment read instead of guessing one).")

    most_posted_theme = d["Theme"].value_counts().idxmax()
    best_avg_theme = d.groupby("Theme")["Engagement"].mean().idxmax()
    pos_share = _pct((d["Sentiment"] == "Positive").sum(), len(d))
    same_theme = best_avg_theme == most_posted_theme
    takeaways = [
        ("\U0001F3C6", f"Their single best-performing post ({int(best['Engagement']):,} engagement, "
                        f"{best['_dt'].strftime('%b %d, %Y')}) was a <b>{_safe(best['Theme'])}</b> post."),
        ("\U0001F4CA", f"<b>{best_avg_theme}</b> drives the most engagement per post on average" +
                        (", and it's also their most-posted theme." if same_theme else
                         f", even though <b>{most_posted_theme}</b> is what they post about most.")),
        ("\U0001F642", f"{pos_share} of their own posts this period read as positive in tone."),
        ("\U0001F4C5", f"Posting cadence is steady at ~{cadence:.1f} posts/week over the window covered."),
    ]
    st.markdown("<div class='takeaway-box'>" + "".join(
        f"<div class='takeaway-item'><span class='takeaway-ic'>{ic}</span><span>{txt}</span></div>"
        for ic, txt in takeaways
    ) + "</div>", unsafe_allow_html=True)
    st.write("")


def render_competitor_profile(name):
    row = _competitor_row(name)
    if row is None:
        st.info(f"No signals found this period for {name}.")
        return

    level = _safe(row.get("Overall activity level"), "n/a")
    level_color = ACTIVITY_LEVEL_COLORS.get(level, GRAY)

    hcol1, hcol2 = st.columns([1, 7])
    with hcol1:
        st.markdown(logo_badge(name, LOGO_MAP, 72), unsafe_allow_html=True)
    with hcol2:
        st.markdown(f"### {name}")
        st.markdown(chip(level + " activity", level_color), unsafe_allow_html=True)
        st.caption(
            f"{int(row['Total activity'])} tracked signals this period across LinkedIn, Instagram, and YouTube."
        )

    st.write("")
    k1, k2, k3, k4 = st.columns(4)
    kpi_card(k1, "\U0001F4CA", "Total activity", int(row["Total activity"]))
    kpi_card(k2, "\U0001F4BC", "LinkedIn posts", int(row.get("LinkedIn posts", 0) or 0))
    kpi_card(k3, "\U0001F4F7", "Instagram posts", int(row.get("Instagram posts", 0) or 0))
    kpi_card(k4, "▶️", "YouTube videos", int(row.get("YouTube videos", 0) or 0))

    st.markdown("##### Where they're active")
    cc1, cc2 = st.columns([3, 2])
    with cc1:
        chan_row = comp_x_chan.loc[comp_x_chan["Competitor"] == name]
        if chan_row.empty:
            st.info("No signals found this period for LinkedIn, Instagram, or YouTube activity.")
        else:
            cr = chan_row.iloc[0]
            chans = {"LinkedIn": cr.get("LinkedIn posts", 0), "Instagram": cr.get("Instagram posts", 0),
                      "YouTube": cr.get("YouTube videos", 0)}
            cdf = pd.DataFrame({"Channel": list(chans.keys()), "Count": [int(v or 0) for v in chans.values()]})
            cdf = cdf[cdf["Count"] > 0]
            if cdf.empty:
                st.info("No signals found this period for LinkedIn, Instagram, or YouTube activity.")
            else:
                cdf = highlight_leader(cdf.sort_values("Count"), "Count")
                fig = px.bar(cdf, x="Count", y="Channel", orientation="h", text="Count", color="_Highlight",
                             color_discrete_map={"Leading": ATLAS_BLUE, "Other": NEUTRAL_BAR})
                fig.update_yaxes(title="")
                fig.update_traces(textposition="outside", cliponaxis=False, textfont=dict(size=11))
                clean_value_axis(fig, "h")
                st.plotly_chart(style_layout(fig, 230, legend="none"), width="stretch")
    with cc2:
        st.markdown("**Top themes**")
        if theme_by_comp.empty or name not in theme_by_comp["Competitor"].values:
            st.info("No signals found this period for themes.")
        else:
            trow = theme_by_comp.loc[theme_by_comp["Competitor"] == name].iloc[0]
            tcols = [c for c in theme_by_comp.columns if c != "Competitor"]
            tvals = trow[tcols]
            tvals = tvals[tvals > 0].sort_values(ascending=False).head(5)
            if tvals.empty:
                st.info("No signals found this period for themes.")
            else:
                pills = "".join(f"<span class='theme-pill'>{t} · {int(c)}</span>" for t, c in tvals.items())
                st.markdown(pills, unsafe_allow_html=True)

    st.write("")
    owned_for_chart = pd.DataFrame()
    if not raw_data.empty:
        _comp_posts = raw_data[(raw_data["Competitor"] == name) & (raw_data["Channel"] != "Jobs")]
        owned_for_chart = _comp_posts[_comp_posts["Source type"].isin(OWNED_SOURCE_TYPES)]
    render_engagement_section(name, owned_for_chart)

    st.markdown("##### Hiring & expansion signals")
    h = hiring[hiring["Competitor"] == name] if not hiring.empty else pd.DataFrame()
    if info_if_empty(h, f"hiring & expansion at {name}"):
        pass
    else:
        st.dataframe(
            h[["Signal type", "Location (HQ)", "Top open-role function (latest snapshot)", "Date", "URL",
               "Signal detail", "Possible meaning"]],
            width="stretch", hide_index=True,
            column_config={"URL": st.column_config.LinkColumn("URL")},
        )

    st.markdown("##### PR, news, events & webinars")
    pr = pr_news[pr_news["Competitor"] == name] if not pr_news.empty else pd.DataFrame()
    if info_if_empty(pr, f"PR, news, events & webinars for {name}"):
        pass
    else:
        st.dataframe(
            pr[["Source channel", "Type", "Title", "Date", "URL", "Theme", "Possible meaning"]],
            width="stretch", hide_index=True,
            column_config={"URL": st.column_config.LinkColumn("URL")},
        )

    st.markdown("##### Product & service signals")
    ps = product_signals[product_signals["Competitor"] == name] if not product_signals.empty else pd.DataFrame()
    if info_if_empty(ps, f"product & service signals for {name}"):
        pass
    else:
        st.dataframe(
            ps[["Product/service category", "Page/post title", "URL", "Date found/published", "Possible meaning"]],
            width="stretch", hide_index=True,
            column_config={"URL": st.column_config.LinkColumn("URL")},
        )

    st.markdown("##### Opportunities for review")
    matches = pd.DataFrame()
    if not opportunities.empty:
        matches = opportunities[opportunities["Opportunity/Signal"].str.contains(name, case=False, na=False)]
    if matches.empty:
        st.caption(f"No opportunity specifically flagged for {name} this period — see the Executive Summary "
                   "for cross-competitor signals.")
    else:
        for _, r in matches.iterrows():
            p_color = PRIORITY_COLORS.get(r["Priority"], GRAY)
            c_color = CONFIDENCE_COLORS.get(r["Confidence"], GRAY)
            icon = PRIORITY_ICONS.get(r["Priority"], "⚪")
            with st.container(border=True):
                st.markdown(f"**{icon} {r['Opportunity/Signal']}**")
                st.markdown(
                    chip(f"Priority: {r['Priority']}", p_color) + chip(f"Confidence: {r['Confidence']}", c_color) +
                    f"<span style='color:{GRAY};font-size:0.82rem;'> Suggested team: {r['Suggested team to review']}</span>",
                    unsafe_allow_html=True,
                )
                st.markdown(f"*Why it may matter:* {r['Why it may matter']}")
                with st.expander("Evidence"):
                    st.write(r["Evidence"])

    st.markdown(
        "##### Recent posts & signals"
        "<span class='section-hint'>· their own channels, plus third-party mentions</span>",
        unsafe_allow_html=True,
    )
    all_posts = raw_data[raw_data["Competitor"] == name].copy() if not raw_data.empty else pd.DataFrame()
    all_posts = all_posts[all_posts["Channel"] != "Jobs"] if not all_posts.empty else all_posts
    if all_posts.empty:
        st.info("No signals found this period.")
    else:
        all_posts["_dt"] = pd.to_datetime(all_posts["Date"], errors="coerce")

        # Coverage strip: one glance at what showed up where before opening anything -
        # lit up in blue when there's content, muted gray when there's none.
        strip_html = ""
        for label, types, icon in OWNED_CHANNEL_BUCKETS + MENTION_TYPE_BUCKETS:
            n = int(all_posts["Source type"].isin(types).sum())
            strip_html += chip(f"{icon} {label} · {n}", ATLAS_BLUE if n else NEUTRAL_BAR, fg=WHITE if n else NAVY)
        st.markdown(strip_html, unsafe_allow_html=True)
        st.write("")

        owned_total = int(all_posts["Source type"].isin(OWNED_SOURCE_TYPES).sum())
        with st.expander(f"\U0001F4E4 Posted by {name} directly — {owned_total} posts across their own channels",
                          expanded=False):
            st.caption("Their own channels - shown separately by platform, not merged into one feed.")
            any_owned = False
            for chan_label, types, icon in OWNED_CHANNEL_BUCKETS:
                sub = all_posts[all_posts["Source type"].isin(types)]
                if sub.empty:
                    continue
                any_owned = True
                st.markdown(f"**{icon} {chan_label}**")
                render_post_section(sub)
                st.write("")
            if not any_owned:
                st.caption(f"No posts published directly by {name} were found this period.")

        mention_total = int(all_posts["Source type"].isin(MENTION_SOURCE_TYPES).sum())
        with st.expander(f"\U0001F4AC Mentions of {name} elsewhere — {mention_total} posts/comments",
                          expanded=False):
            st.caption("Other people's posts and comments using their name - shown separately by type.")
            any_mention = False
            for mention_label, types, icon in MENTION_TYPE_BUCKETS:
                sub = all_posts[all_posts["Source type"].isin(types)]
                if sub.empty:
                    continue
                any_mention = True
                st.markdown(f"**{icon} {mention_label}**")
                render_post_section(sub)
                st.write("")
            if not any_mention:
                st.caption(f"No third-party mentions of {name} were found this period.")

        covered_types = OWNED_SOURCE_TYPES | MENTION_SOURCE_TYPES
        leftover = all_posts[~all_posts["Source type"].isin(covered_types)]
        if not leftover.empty:
            with st.expander(f"\U0001F9E9 Other signals — {len(leftover)}", expanded=False):
                render_post_section(leftover)


def render_other_competitors():
    st.subheader("Other Competitors")
    st.caption("Lower overall activity this period, but still worth a quick scan. Use the Raw Data Explorer tab "
               "to dig into any of them in full detail.")
    if not OTHER_COMPETITORS:
        st.info("All tracked competitors are covered in their own tab.")
        return

    rows = []
    for nm in OTHER_COMPETITORS:
        r = _competitor_row(nm)
        if r is None:
            continue
        rows.append({
            "Logo": LOGO_MAP.get(nm),
            "Competitor": nm,
            "Activity level": _safe(r.get("Overall activity level"), "n/a"),
            "Total activity": int(r["Total activity"]),
            "LinkedIn": int(r.get("LinkedIn posts", 0) or 0),
            "Instagram": int(r.get("Instagram posts", 0) or 0),
            "YouTube": int(r.get("YouTube videos", 0) or 0),
            "Top channel": _top_channel_for(r),
            "Top theme": _top_theme_for(nm),
        })
    odf = pd.DataFrame(rows).sort_values("Total activity", ascending=False)

    if not odf.empty:
        leader = odf.iloc[0]
        insight(
            f"Among the remaining tracked competitors, <b>{leader['Competitor']}</b> is the most active with "
            f"{leader['Total activity']} signals this period — still well below the top 4 shown in their own tabs."
        )

    st.dataframe(
        odf, width="stretch", hide_index=True,
        column_config={
            "Logo": st.column_config.ImageColumn("Logo"),
        },
    )
    st.caption("Logos shown only where a source image was available in the scraped data — no logo is shown "
               "rather than guessing.")


# ---------------------------------------------------------------------------
# Header + disclaimer
# ---------------------------------------------------------------------------
st.markdown(
    f"<div class='main-header'><div class='main-header-icon'>\U0001F9ED</div>"
    f"<div><div class='main-header-title'>Competitor Intelligence Dashboard</div>"
    f"<div class='main-header-sub'>Air Compressor Category &middot; "
    f"{settings.get('Baseline_Month','May 2026')} vs {settings.get('Report_Month','June 2026')}</div>"
    f"</div></div>",
    unsafe_allow_html=True,
)

_tab_labels = ["1. Executive Summary"]
for _i, _name in enumerate(TOP_COMPETITORS):
    _tab_labels.append(f"{_i + 2}. {_name}")
_tab_labels.append(f"{len(TOP_COMPETITORS) + 2}. Other Competitors")
_tab_labels.append(f"{len(TOP_COMPETITORS) + 3}. Raw Data Explorer")
tabs = st.tabs(_tab_labels)

# ===========================================================================
# TAB 1 - EXECUTIVE SUMMARY  ("the killer one")
# ===========================================================================
with tabs[0]:

    # ---- The headline story (exact required template sentence) -----------
    st.markdown(
        f"<div class='story-banner'>\U0001F4F0 <b>The story this month:</b><br>{exec_metric('Executive summary narrative')}</div>",
        unsafe_allow_html=True,
    )

    # ---- Plain-English elaboration, computed live from the data ----------
    elab_bits = []
    if not overview.empty:
        ov = overview.sort_values("Total activity", ascending=False).reset_index(drop=True)
        leader = ov.iloc[0]
        if len(ov) > 1 and ov.iloc[1]["Total activity"]:
            second = ov.iloc[1]
            gap_pct = (leader["Total activity"] - second["Total activity"]) / second["Total activity"] * 100
            elab_bits.append(
                f"<b>{leader['Competitor']}</b> was the single most active competitor, with "
                f"{int(leader['Total activity'])} tracked signals - roughly {gap_pct:.0f}% more than "
                f"<b>{second['Competitor']}</b>, the next-busiest brand."
            )
    if not channel_df.empty:
        cd = channel_df.sort_values("Total activity (May+June 2026)", ascending=False).reset_index(drop=True)
        total_chan = cd["Total activity (May+June 2026)"].sum()
        top_chan = cd.iloc[0]
        elab_bits.append(
            f"<b>{top_chan['Channel']}</b> carried the bulk of the conversation, making up "
            f"{_pct(top_chan['Total activity (May+June 2026)'], total_chan)} of everything we tracked."
        )
    if not conv_sentiment.empty:
        total_sent = conv_sentiment["Count"].sum()
        pos = conv_sentiment.loc[conv_sentiment["Sentiment"] == "Positive", "Count"].sum()
        neg = conv_sentiment.loc[conv_sentiment["Sentiment"] == "Negative", "Count"].sum()
        elab_bits.append(
            f"Public conversation about the category skewed positive ({_pct(pos, total_sent)} positive vs. "
            f"{_pct(neg, total_sent)} negative), so there's no broad reputational alarm bell ringing right now."
        )
    if elab_bits:
        st.markdown(f"<div class='elaborate-box'>{' '.join(elab_bits)}</div>", unsafe_allow_html=True)

    # ---- The Big Picture: 4 synthesis charts -------------------------------
    st.markdown("#### \U0001F50D The big picture")
    cc1, cc2 = st.columns(2)
    with cc1:
        if not overview.empty:
            top_act = overview.sort_values("Total activity", ascending=False)
            fig = px.bar(top_act, x="Total activity", y="Competitor", orientation="h",
                         color="Overall activity level", text="Total activity",
                         color_discrete_map=ACTIVITY_LEVEL_COLORS,
                         title="Who's making the most noise? (activity by competitor)")
            fig.update_yaxes(autorange="reversed", title="")
            fig.update_traces(textposition="outside", cliponaxis=False, textfont=dict(size=11))
            clean_value_axis(fig, "h")
            st.plotly_chart(style_layout(fig, 400, legend="bottom"), width="stretch")
            st.caption("\U0001F4A1 Bigger bars = more posts, jobs, and content found from that competitor this period.")
    with cc2:
        if not channel_df.empty:
            fig = px.pie(channel_df, names="Channel", values="Total activity (May+June 2026)", hole=0.5,
                         color_discrete_sequence=CATEGORICAL_PALETTE, title="Where is the activity happening?")
            fig.update_traces(textinfo="percent", textfont_size=11)
            st.plotly_chart(style_layout(fig, 400, legend="right"), width="stretch")
            st.caption("\U0001F4A1 Shows which channel (LinkedIn, Instagram, etc.) competitors are using most.")

    cc3, cc4 = st.columns(2)
    with cc3:
        if not theme_freq.empty:
            tf = theme_freq[theme_freq["Theme"] != "General brand activity"].sort_values("Frequency", ascending=True).tail(8)
            tf = highlight_leader(tf, "Frequency")
            fig = px.bar(tf, x="Frequency", y="Theme", orientation="h", text="Frequency", color="_Highlight",
                         color_discrete_map={"Leading": ATLAS_BLUE, "Other": NEUTRAL_BAR},
                         title="What are competitors talking about?")
            fig.update_yaxes(title="")
            fig.update_traces(textposition="outside", cliponaxis=False, textfont=dict(size=11))
            clean_value_axis(fig, "h")
            st.plotly_chart(style_layout(fig, 400, legend="none"), width="stretch")
            st.caption("\U0001F4A1 The most common subjects competitors are posting and writing about.")
    with cc4:
        if not conv_sentiment.empty:
            fig = px.pie(conv_sentiment, names="Sentiment", values="Count", hole=0.5,
                         color="Sentiment", color_discrete_map=SENTIMENT_COLORS, title="How is the market reacting?")
            fig.update_traces(textinfo="percent", textfont_size=11)
            st.plotly_chart(style_layout(fig, 400, legend="right"), width="stretch")
            st.caption("\U0001F4A1 Tone of comments and posts mentioning brands in the category - mostly neutral-to-positive is a healthy sign.")

    # ---- Signals snapshot: smaller data points, compact cards --------------
    st.markdown("#### \U0001F9E9 Other signals worth knowing about")
    s1, s2, s3, s4 = st.columns(4)
    if not brand_mentions.empty:
        bm = brand_mentions[~brand_mentions["Brand"].isin(["Unspecified / General"])].sort_values("Mentions", ascending=False)
        if not bm.empty:
            top_bm = bm.iloc[0]
            mini_stat(s1, "Most-talked-about brand", f"{top_bm['Brand']} — {int(top_bm['Mentions'])} mentions in social conversation")
    if not hiring.empty:
        hc = hiring["Competitor"].value_counts()
        mini_stat(s2, "Most active hirer", f"{hc.index[0]} — {int(hc.iloc[0])} of {len(hiring)} hiring/expansion signals")
    zero_channels_exec = []
    for chan_label in ["Events", "Webinars", "PR releases", "News mentions"]:
        if channel_df.empty or chan_label not in channel_df["Channel"].values or \
           channel_df.loc[channel_df["Channel"] == chan_label, "Total activity (May+June 2026)"].sum() == 0:
            zero_channels_exec.append(chan_label)
    if zero_channels_exec:
        mini_stat(s3, "Coverage gap to validate", f"0 confirmed {_join_natural(zero_channels_exec)} this period — worth checking with PR team")
    our_brand_name = str(settings.get("Our_Brand", "Atlas Copco"))
    if not brand_mentions.empty and our_brand_name in brand_mentions["Brand"].values:
        our_mentions = int(brand_mentions.loc[brand_mentions["Brand"] == our_brand_name, "Mentions"].iloc[0])
        mini_stat(s4, "Our brand in the conversation", f"{our_brand_name} mentioned {our_mentions} times this period")

    # ---- Key themes to watch, in plain English -----------------------------
    st.markdown("#### \U0001F3AF Key themes to watch")
    theme_lookup = {}
    if not theme_freq.empty:
        total_themed = theme_freq.loc[theme_freq["Theme"] != "General brand activity", "Frequency"].sum()
        theme_lookup = dict(zip(theme_freq["Theme"], theme_freq["Frequency"]))
    tcols = st.columns(3)
    for i, col in zip((1, 2, 3), tcols):
        theme_name = exec_metric(f"Top theme #{i}")
        freq = theme_lookup.get(theme_name)
        share = f" ({_pct(freq, total_themed)} of themed content)" if freq and theme_lookup else ""
        with col:
            st.markdown(f"<span class='theme-pill'>{theme_name}</span>{share}", unsafe_allow_html=True)

    # ---- Top opportunities, as compact cards --------------------------------
    st.markdown("#### ⭐ Top opportunities for review")
    ocols = st.columns(3)
    for i, col in zip((1, 2, 3), ocols):
        opp_text = exec_metric(f"Possible opportunity #{i}")
        with col:
            st.markdown(
                f"<div class='next-step'>{PRIORITY_ICONS.get('High','⭐')} {opp_text}</div>",
                unsafe_allow_html=True,
            )

    # ---- What to do next: numbered action cards -----------------------------
    st.markdown("#### ✅ What to do next")
    if not opportunities.empty:
        prio_rank = {"High": 0, "Medium": 1, "Low": 2}
        top_opps = opportunities.copy()
        top_opps["_r"] = top_opps["Priority"].map(prio_rank).fillna(3)
        top_opps = top_opps.sort_values("_r").head(5).reset_index(drop=True)
        for i, row in top_opps.iterrows():
            teams = [t.strip() for t in str(row["Suggested team to review"]).split(";") if t.strip()]
            team_chips = "".join(chip(t, NAVY) for t in teams)
            p_icon = PRIORITY_ICONS.get(row["Priority"], "⭐")
            with st.container(border=True):
                ncol1, ncol2 = st.columns([1, 14])
                with ncol1:
                    st.markdown(f"##### {i + 1}.")
                with ncol2:
                    st.markdown(f"**{p_icon} {row['Opportunity/Signal']}**")
                    st.caption(row["Why it may matter"])
                    if team_chips:
                        st.markdown(team_chips, unsafe_allow_html=True)
    st.caption("Each top competitor's own tab has the full opportunity detail and evidence behind items related to them.")

    st.markdown("---")
    st.caption(DISCLAIMER)

# ===========================================================================
# TABS 2-5 - ONE PROFILE TAB PER TOP-ACTIVITY COMPETITOR (ranked live)
# ===========================================================================
for _i, _name in enumerate(TOP_COMPETITORS):
    with tabs[1 + _i]:
        render_competitor_profile(_name)

# ===========================================================================
# OTHER COMPETITORS TAB
# ===========================================================================
with tabs[1 + len(TOP_COMPETITORS)]:
    render_other_competitors()

# ===========================================================================
# RAW DATA EXPLORER TAB
# ===========================================================================
with tabs[2 + len(TOP_COMPETITORS)]:
    st.subheader("Raw Data Explorer")
    st.caption("Every record behind this dashboard. Use the filters below to search and narrow down; export "
               "a filtered copy with the download button.")

    if not info_if_empty(raw_data, "raw data"):
        f1, f2, f3, f4, f5 = st.columns(5)
        month_pick = f1.multiselect("Month", sorted(raw_data["Month"].dropna().unique()), key="raw_month")
        comp_pick = f2.multiselect("Competitor", sorted(raw_data["Competitor"].dropna().unique()), key="raw_comp")
        chan_pick = f3.multiselect("Channel", sorted(raw_data["Channel"].dropna().unique()), key="raw_chan")
        theme_pick = f4.multiselect("Theme", sorted(raw_data["Theme"].dropna().unique()), key="raw_theme")
        team_pick = f5.multiselect("Team to review", sorted(raw_data["Team to review"].dropna().unique()), key="raw_team")

        f6, f7, f8, f9, f10 = st.columns(5)
        type_pick = f6.multiselect("Source type", sorted(raw_data["Source type"].dropna().unique()), key="raw_type")
        prio_pick = f7.multiselect("Priority", sorted(raw_data["Priority"].dropna().unique()), key="raw_prio")
        sent_pick = f8.multiselect("Sentiment", sorted(raw_data["Sentiment"].dropna().unique()), key="raw_sent")
        geo_options = sorted(set(g.strip() for gs in raw_data["Geography"].dropna() for g in str(gs).split(";") if g.strip()))
        geo_pick = f9.multiselect("Region / location", geo_options, key="raw_geo")
        cat_options = sorted(raw_data["Product/service category"].dropna().unique())
        cat_pick = f10.multiselect("Product/service category", cat_options, key="raw_cat")

        search = st.text_input("Search title / summary", "", key="raw_search")

        view = raw_data.copy()
        if month_pick:
            view = view[view["Month"].isin(month_pick)]
        if comp_pick:
            view = view[view["Competitor"].isin(comp_pick)]
        if chan_pick:
            view = view[view["Channel"].isin(chan_pick)]
        if theme_pick:
            view = view[view["Theme"].isin(theme_pick)]
        if team_pick:
            view = view[view["Team to review"].isin(team_pick)]
        if type_pick:
            view = view[view["Source type"].isin(type_pick)]
        if prio_pick:
            view = view[view["Priority"].isin(prio_pick)]
        if sent_pick:
            view = view[view["Sentiment"].isin(sent_pick)]
        if geo_pick:
            view = view[view["Geography"].apply(lambda g: any(loc in str(g) for loc in geo_pick))]
        if cat_pick:
            view = view[view["Product/service category"].isin(cat_pick)]
        if search:
            mask = (view["Title"].str.contains(search, case=False, na=False) |
                    view["Summary"].str.contains(search, case=False, na=False))
            view = view[mask]

        st.caption(f"Showing {len(view):,} of {len(raw_data):,} records")
        st.dataframe(
            priority_badge_table(view),
            width="stretch", hide_index=True, height=520,
            column_config={"URL": st.column_config.LinkColumn("URL")},
        )
        st.download_button(
            "Download filtered data as CSV",
            data=view.to_csv(index=False).encode("utf-8"),
            file_name="competitor_intelligence_filtered.csv",
            mime="text/csv",
        )
