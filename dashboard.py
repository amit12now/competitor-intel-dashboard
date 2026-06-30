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
import re
from html import escape
from urllib.parse import parse_qs, urlparse

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
    "Clean_June_Competitor_Brief", "Clean_June_Content_Mix",
    "Clean_June_Social_Metrics", "Clean_June_Content_Items",
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
div[data-testid="stMarkdown"] h5 {{
    background:{WHITE};
    border:1px solid #DDE5EC;
    border-left:5px solid {ATLAS_BLUE};
    border-radius:8px;
    padding:13px 16px;
    margin:30px 0 16px 0;
    color:{NAVY};
    box-shadow:0 1px 4px rgba(15, 23, 42, 0.035);
}}
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
.kpi-card, .mini-stat, .next-step, .theme-rank-card, .opp-card {{ transition: box-shadow 0.15s ease; }}
.kpi-card:hover, .mini-stat:hover, .theme-rank-card:hover, .opp-card:hover {{ box-shadow: 0 4px 12px rgba(0,0,0,0.08); }}
.theme-rank-card {{
    background-color: {WHITE};
    border: 1px solid #E2E8F0;
    border-radius: 10px;
    padding: 14px 16px;
    height: 100%;
}}
.theme-rank-top {{ display: flex; align-items: center; gap: 10px; margin-bottom: 6px; }}
.theme-rank-badge {{
    width: 24px; height: 24px; border-radius: 50%;
    background: {ATLAS_BLUE}; color: {WHITE};
    display: flex; align-items: center; justify-content: center;
    font-size: 0.74rem; font-weight: 800; flex: 0 0 auto;
}}
.theme-rank-name {{ color: {NAVY}; font-size: 0.88rem; font-weight: 700; }}
.theme-rank-pct {{ color: {ATLAS_BLUE_DARK}; font-size: 1.5rem; font-weight: 800; margin-bottom: 8px; }}
.theme-rank-track {{ background: {GRAY_LIGHT}; border-radius: 6px; height: 8px; overflow: hidden; }}
.theme-rank-fill {{ background: linear-gradient(90deg, {ATLAS_BLUE} 0%, {ATLAS_BLUE_DARK} 100%); height: 100%; border-radius: 6px; }}
.opp-card {{
    background-color: {WHITE};
    border: 1px solid #E2E8F0;
    border-top: 4px solid {ATLAS_BLUE};
    border-radius: 10px;
    padding: 14px 16px;
    height: 100%;
    font-size: 0.88rem;
    color: {NAVY};
    line-height: 1.4;
}}
.opp-card-title {{ font-weight: 700; }}
.action-card {{
    display: flex; gap: 14px; align-items: flex-start;
    background-color: {WHITE};
    border: 1px solid #E2E8F0;
    border-left: 4px solid {ATLAS_BLUE};
    border-radius: 10px;
    padding: 14px 18px;
    margin-bottom: 10px;
    transition: box-shadow 0.15s ease;
}}
.action-card:hover {{ box-shadow: 0 4px 12px rgba(0,0,0,0.08); }}
.action-num {{
    flex: 0 0 auto;
    width: 28px; height: 28px; border-radius: 50%;
    background: {NAVY}; color: {WHITE};
    display: flex; align-items: center; justify-content: center;
    font-weight: 800; font-size: 0.88rem; margin-top: 1px;
}}
.action-title {{ color: {NAVY}; font-size: 0.92rem; font-weight: 700; margin-bottom: 4px; line-height: 1.4; }}
.action-why {{ color: {GRAY}; font-size: 0.84rem; line-height: 1.5; margin-bottom: 6px; }}
.card-meta {{
    display: flex; align-items: center; flex-wrap: wrap; gap: 6px;
    color: {GRAY}; font-size: 0.78rem; font-weight: 600; margin-top: 8px;
}}
.meta-dot {{ width: 7px; height: 7px; border-radius: 50%; display: inline-block; flex: 0 0 auto; }}
.team-tag {{
    display: inline-block; font-size: 0.72rem; font-weight: 600;
    padding: 2px 9px; border-radius: 6px;
    background: {GRAY_LIGHT}; color: {NAVY}; border: 1px solid #E2E8F0;
}}
.brief-panel {{
    background: {WHITE};
    border: 1px solid #DDE5EC;
    border-radius: 10px;
    padding: 18px 20px 20px 20px;
    margin: 12px 0 18px 0;
    box-shadow: 0 1px 4px rgba(15, 23, 42, 0.04);
}}
.exec-divider {{
    height: 1px;
    background: linear-gradient(90deg, rgba(0,153,204,0.42), rgba(221,229,236,0.96), rgba(221,229,236,0));
    margin: 34px 0 22px 0;
}}
.brief-head {{
    display: flex; justify-content: space-between; align-items: flex-start;
    gap: 18px;
    background: {WHITE};
    border: 1px solid #DDE5EC;
    border-left: 5px solid {ATLAS_BLUE};
    border-radius: 8px;
    padding: 14px 16px 13px 16px;
    margin: 2px 0 18px 0;
    box-shadow: 0 1px 4px rgba(15, 23, 42, 0.035);
}}
.brief-eyebrow {{
    color: {ATLAS_BLUE_DARK}; font-size: 0.72rem; font-weight: 800;
    text-transform: uppercase; letter-spacing: 0.06em; margin-bottom: 5px;
}}
.brief-title {{ color: {NAVY}; font-size: 1.18rem; font-weight: 800; line-height: 1.25; }}
.brief-sub {{ color: {GRAY}; font-size: 0.88rem; line-height: 1.52; margin-top: 5px; max-width: 940px; }}
.brief-subhead {{
    background: #F8FAFC;
    border-left-color: #9AA5B1;
    box-shadow: none;
    margin-top: 22px;
}}
.brief-subhead .brief-title {{ font-size: 1.02rem; }}
.brief-subhead .brief-sub {{ font-size: 0.82rem; }}
.signal-card {{
    background: {WHITE};
    border: 1px solid #E2E8F0;
    border-radius: 8px;
    padding: 14px 15px;
    height: 100%;
    box-shadow: 0 1px 3px rgba(15, 23, 42, 0.04);
}}
.signal-top {{ display:flex; justify-content:space-between; gap:10px; align-items:flex-start; }}
.signal-label {{ color:{GRAY}; font-size:0.72rem; font-weight:800; text-transform:uppercase; letter-spacing:0.04em; }}
.signal-value {{ color:{NAVY}; font-size:1.25rem; font-weight:800; line-height:1.2; margin-top:5px; }}
.signal-note {{ color:{GRAY}; font-size:0.80rem; line-height:1.42; margin-top:7px; }}
.signal-rule {{ width:30px; height:4px; border-radius:4px; background:{ATLAS_BLUE}; margin-top:3px; flex:0 0 auto; }}
.rank-card {{
    background:{WHITE};
    border:1px solid #E2E8F0;
    border-radius:8px;
    padding:14px 15px;
    min-height:150px;
    box-shadow:0 1px 3px rgba(15,23,42,0.04);
}}
.rank-top {{ display:flex; gap:10px; align-items:center; margin-bottom:8px; }}
.rank-num {{
    width:26px; height:26px; border-radius:6px;
    background:{ATLAS_BLUE_LIGHT}; color:{ATLAS_BLUE_DARK};
    display:flex; align-items:center; justify-content:center;
    font-weight:800; font-size:0.78rem; flex:0 0 auto;
}}
.rank-title {{ color:{NAVY}; font-size:0.92rem; font-weight:800; line-height:1.3; }}
.rank-metric {{ color:{ATLAS_BLUE_DARK}; font-size:1.35rem; font-weight:800; margin:7px 0 5px; }}
.rank-copy {{ color:{GRAY}; font-size:0.80rem; line-height:1.45; }}
.social-card {{
    background:{WHITE}; border:1px solid #E2E8F0; border-radius:8px;
    padding:14px 15px; height:100%; box-shadow:0 1px 3px rgba(15,23,42,0.04);
}}
.social-card-title {{ color:{NAVY}; font-size:0.90rem; font-weight:800; line-height:1.35; margin-bottom:8px; }}
.social-card-meta {{ color:{GRAY}; font-size:0.78rem; font-weight:700; margin-bottom:8px; }}
.social-card-kpis {{ display:flex; flex-wrap:wrap; gap:7px; margin-top:8px; }}
.social-kpi {{ background:{GRAY_LIGHT}; color:{NAVY}; border:1px solid #E2E8F0; border-radius:6px; padding:3px 8px; font-size:0.72rem; font-weight:700; }}
.news-card {{
    background:{WHITE}; border:1px solid #E2E8F0; border-radius:8px;
    padding:13px 14px; height:100%; box-shadow:0 1px 3px rgba(15,23,42,0.04);
}}
.news-title {{ color:{NAVY}; font-size:0.86rem; font-weight:800; line-height:1.35; }}
.news-meta {{ color:{GRAY}; font-size:0.76rem; margin-top:7px; line-height:1.4; }}
.source-card {{
    background:{WHITE};
    border:1px solid #E2E8F0;
    border-left:4px solid {ATLAS_BLUE};
    border-radius:8px;
    padding:14px 15px;
    min-height:155px;
    height:100%;
    box-shadow:0 1px 3px rgba(15,23,42,0.04);
}}
.source-eyebrow {{
    color:{ATLAS_BLUE_DARK};
    font-size:0.70rem;
    font-weight:800;
    text-transform:uppercase;
    letter-spacing:0.05em;
    margin-bottom:7px;
}}
.source-title {{
    color:{NAVY};
    font-size:0.92rem;
    font-weight:800;
    line-height:1.36;
    margin-bottom:8px;
}}
.source-note {{ color:{GRAY}; font-size:0.80rem; line-height:1.48; margin-top:8px; }}
.source-link {{ color:{ATLAS_BLUE_DARK}; font-size:0.78rem; font-weight:800; margin-top:10px; }}
.source-link a {{ color:{ATLAS_BLUE_DARK}; text-decoration:none; }}
.post-card {{
    background:{WHITE};
    border:1px solid #E2E8F0;
    border-radius:8px;
    overflow:hidden;
    height:100%;
    box-shadow:0 1px 3px rgba(15,23,42,0.04);
}}
.post-card-body {{ padding:13px 14px 14px 14px; }}
.post-card-img {{ width:100%; aspect-ratio:16/9; object-fit:cover; display:block; background:{GRAY_LIGHT}; }}
.post-card-placeholder {{
    width:100%;
    aspect-ratio:16/9;
    display:flex;
    flex-direction:column;
    justify-content:flex-end;
    padding:16px;
    background:linear-gradient(135deg, {ATLAS_BLUE_LIGHT} 0%, #F8FAFC 58%, #EEF2F6 100%);
    border-bottom:1px solid #E2E8F0;
}}
.post-card-placeholder .platform {{
    color:{ATLAS_BLUE_DARK};
    font-size:0.72rem;
    font-weight:900;
    text-transform:uppercase;
    letter-spacing:0.06em;
}}
.post-card-placeholder .bucket {{
    color:{NAVY};
    font-size:1.02rem;
    font-weight:900;
    line-height:1.2;
    margin-top:4px;
}}
.post-card-placeholder .note {{
    color:{GRAY};
    font-size:0.72rem;
    font-weight:700;
    margin-top:6px;
}}
.post-card-title {{ color:{NAVY}; font-size:0.94rem; font-weight:800; line-height:1.38; margin:8px 0; }}
.post-card-meta {{ color:{GRAY}; font-size:0.76rem; font-weight:700; line-height:1.4; margin-top:6px; }}
.post-card-score {{
    display:inline-flex; align-items:center; gap:5px;
    background:{ATLAS_BLUE_LIGHT}; color:{ATLAS_BLUE_DARK};
    border:1px solid #B9E4F2; border-radius:6px;
    padding:4px 9px; font-size:0.75rem; font-weight:800;
}}
.post-card-metrics {{ display:flex; flex-wrap:wrap; gap:6px; margin-top:10px; }}
.platform-filter-label {{
    color:{GRAY};
    font-size:0.78rem;
    font-weight:800;
    text-transform:uppercase;
    letter-spacing:0.04em;
    margin:16px 0 5px;
}}
.snapshot-table {{ border:1px solid #DDE5EC; border-radius:8px; overflow:hidden; background:{WHITE}; }}
.snapshot-row {{
    display:grid; grid-template-columns: 46px minmax(145px, 1.1fr) minmax(120px, .8fr) minmax(120px, .8fr) minmax(150px, 1fr);
    gap:12px; align-items:center; padding:11px 13px; border-bottom:1px solid #E8EEF3;
}}
.snapshot-row:last-child {{ border-bottom:none; }}
.snapshot-head {{ background:#F8FAFC; color:{GRAY}; font-size:0.72rem; font-weight:800; text-transform:uppercase; letter-spacing:0.04em; }}
.snapshot-rank {{ color:{ATLAS_BLUE_DARK}; font-weight:800; }}
.snapshot-name {{ color:{NAVY}; font-weight:800; font-size:0.88rem; }}
.snapshot-small {{ color:{GRAY}; font-size:0.76rem; font-weight:700; }}
.snapshot-bar {{ height:7px; background:{GRAY_LIGHT}; border-radius:8px; overflow:hidden; margin-top:4px; }}
.snapshot-fill {{ height:100%; background:{ATLAS_BLUE}; border-radius:8px; }}
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


def _display_date(value):
    dt = pd.to_datetime(value, errors="coerce")
    return dt.strftime("%b %d, %Y") if pd.notna(dt) else _safe(value)


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


def brief_header(title, subtitle="", eyebrow="Executive view", compact=False):
    sub_html = f"<div class='brief-sub'>{subtitle}</div>" if subtitle else ""
    class_name = "brief-head brief-subhead" if compact else "brief-head"
    st.markdown(
        f"<div class='{class_name}'><div><div class='brief-eyebrow'>{eyebrow}</div>"
        f"<div class='brief-title'>{title}</div>{sub_html}</div>"
        f"<div class='signal-rule'></div></div>",
        unsafe_allow_html=True,
    )


def section_divider():
    st.markdown("<div class='exec-divider'></div>", unsafe_allow_html=True)


def signal_card(col, label, value, note="", accent=ATLAS_BLUE):
    note_html = f"<div class='signal-note'>{note}</div>" if note else ""
    col.markdown(
        f"<div class='signal-card'><div class='signal-top'>"
        f"<div><div class='signal-label'>{label}</div><div class='signal-value'>{value}</div></div>"
        f"<div class='signal-rule' style='background:{accent};'></div></div>{note_html}</div>",
        unsafe_allow_html=True,
    )


def source_card(col, eyebrow, title, meta_bits=None, note="", url="", accent=ATLAS_BLUE, chips_html=""):
    meta_bits = [escape(_safe(m)) for m in (meta_bits or []) if _safe(m)]
    meta_html = f"<div class='card-meta'>{' &middot; '.join(meta_bits)}</div>" if meta_bits else ""
    note_html = f"<div class='source-note'>{escape(_safe(note))}</div>" if _safe(note) else ""
    link_html = (
        f"<div class='source-link'><a href='{escape(_safe(url))}' target='_blank'>View source &rarr;</a></div>"
        if _safe(url) else ""
    )
    col.markdown(
        f"<div class='source-card' style='border-left-color:{accent};'>"
        f"<div class='source-eyebrow'>{escape(_safe(eyebrow))}</div>"
        f"<div class='source-title'>{escape(_safe(title, 'Untitled signal'))}</div>"
        f"{chips_html}{meta_html}{note_html}{link_html}</div>",
        unsafe_allow_html=True,
    )


def compact_detail_expander(label, df, cols):
    available = [c for c in cols if c in df.columns]
    if not available:
        return
    with st.expander(label):
        st.dataframe(
            df[available].reset_index(drop=True),
            width="stretch",
            hide_index=True,
            column_config={"URL": st.column_config.LinkColumn("URL")} if "URL" in available else None,
        )


def simple_bar(value, max_value, color=ATLAS_BLUE):
    width = 0 if not max_value else min(max(value / max_value * 100, 0), 100)
    return (
        f"<div class='theme-rank-track'><div class='theme-rank-fill' "
        f"style='width:{width:.0f}%;background:{color};'></div></div>"
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
june_brief = sheets["Clean_June_Competitor_Brief"]
june_content_mix = sheets["Clean_June_Content_Mix"]
june_social_metrics = sheets["Clean_June_Social_Metrics"]
june_content_items = sheets["Clean_June_Content_Items"]


def exec_metric(label):
    row = exec_df[exec_df["Metric"] == label]
    return row["Value"].iloc[0] if not row.empty else "n/a"


def _june_brief_row(name):
    if june_brief.empty or "Competitor" not in june_brief.columns:
        return None
    r = june_brief.loc[june_brief["Competitor"] == name]
    return r.iloc[0] if not r.empty else None


def _num(row, col, default=0):
    if row is None:
        return default
    try:
        value = row.get(col, default)
        if pd.isna(value):
            return default
        return int(value)
    except Exception:
        return default


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
        metric_bits = _engagement_breakdown(p)
        if metric_bits:
            st.caption(metric_bits)
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
    ("Posts mentioning 'air compressor'", {"LinkedIn posts mentioning 'air compressor'"}, "\U0001F527"),
]
MENTION_SOURCE_TYPES = {t for _, types, _ in MENTION_TYPE_BUCKETS for t in types}

# ---- Platform tabs: LinkedIn / Instagram / YouTube own-channel feeds, each
# divided by content-type buckets built from the real Theme column. The 15
# granular Theme values are mapped into 4 named buckets with explicit rules -
# nothing invented. Anything that doesn't fit one of the 4 lands in "Other"
# rather than being forced into a bucket it doesn't belong in.
PLATFORM_SOURCE_TYPES = [
    ("LinkedIn", {"LinkedIn company posts", "LinkedIn posts from company employees"}),
    ("Instagram", {"Instagram posts"}),
    ("YouTube", {"YouTube videos"}),
]

THEME_BUCKETS = [
    ("Thought Leadership", {"Sustainability", "Industrial productivity", "Energy efficiency"}),
    ("Product & Equipment", {"New product launch", "Oil-free compressors", "Rotary screw compressors"}),
    ("Webinar", {"Training / webinar"}),
    ("Services", {"Service / aftermarket", "Dealer / distributor"}),
]
THEME_BUCKET_ORDER = [b for b, _ in THEME_BUCKETS] + ["Other"]


def theme_bucket(theme):
    for label, themes in THEME_BUCKETS:
        if theme in themes:
            return label
    return "Other"


def post_format(url):
    """Real, not fabricated: "Image" if we actually have a scraped image for
    this URL (present in IMAGE_MAP), otherwise "Text". There's no native
    format field in the data, so whether a real image asset was captured is
    the closest honest signal available."""
    return "Image" if IMAGE_MAP.get(_safe(url)) else "Text"


def _youtube_thumbnail(url):
    url = _safe(url)
    if not url:
        return ""
    try:
        parsed = urlparse(url)
        host = parsed.netloc.lower()
        if "youtu.be" in host:
            video_id = parsed.path.strip("/").split("/")[0]
        elif "youtube.com" in host:
            qs_id = parse_qs(parsed.query).get("v", [""])[0]
            if qs_id:
                video_id = qs_id
            else:
                match = re.search(r"/(?:shorts|embed|watch)/([^/?#]+)", parsed.path)
                video_id = match.group(1) if match else ""
        else:
            video_id = ""
    except Exception:
        video_id = ""
    return f"https://img.youtube.com/vi/{video_id}/hqdefault.jpg" if video_id else ""


def post_media_html(row, bucket="", source_type=""):
    url = _safe(row.get("URL"))
    img_urls = IMAGE_MAP.get(url, [])
    image_url = _safe(img_urls[0]) if img_urls else ""
    if not image_url:
        image_url = _youtube_thumbnail(url)
    if image_url:
        return f"<img class='post-card-img' src='{escape(image_url)}' />"

    platform = _safe(row.get("Channel")) or _safe(source_type, "Social signal")
    bucket_label = _safe(bucket) or theme_bucket(_safe(row.get("Theme"))) or "Content signal"
    return (
        "<div class='post-card-placeholder'>"
        f"<div class='platform'>{escape(platform)}</div>"
        f"<div class='bucket'>{escape(bucket_label)}</div>"
        "<div class='note'>No source image captured</div>"
        "</div>"
    )


def _metric_chip(label, value):
    try:
        n = int(value or 0)
    except Exception:
        n = 0
    return f"{label} {n:,}" if n else ""


def _engagement_breakdown(row):
    parts = [
        _metric_chip("Likes", row.get("Likes")),
        _metric_chip("Comments", row.get("Comments")),
        _metric_chip("Shares", row.get("Shares")),
        _metric_chip("Views", row.get("Views")),
    ]
    return " · ".join([p for p in parts if p])


def render_platform_post_card(p):
    """Card for the LinkedIn / Instagram / YouTube platform tabs - image,
    content-type bucket, format, date, and real engagement metrics."""
    with st.container(border=True):
        img_urls = IMAGE_MAP.get(_safe(p.get("URL")), [])
        if img_urls:
            try:
                st.image(img_urls[0], width="stretch")
            except Exception:
                pass

        bucket = theme_bucket(_safe(p.get("Theme")))
        fmt = _safe(p.get("Content format")) or post_format(p.get("URL"))
        st.markdown(chip(bucket, ATLAS_BLUE) + chip(fmt, "#6B7280"), unsafe_allow_html=True)

        title = _safe(p.get("Title")) or _safe(p.get("Summary"))[:80]
        if title:
            st.markdown(f"**{title[:90]}**")

        dt = pd.to_datetime(p.get("Date"), errors="coerce")
        date_str = dt.strftime("%b %d, %Y") if pd.notna(dt) else ""
        eng = p.get("Engagement")
        eng_str = f"{int(eng):,} total engagement" if pd.notna(eng) else ""
        meta_bits = [b for b in [date_str, eng_str] if b]
        if meta_bits:
            st.caption(" · ".join(meta_bits))
        metric_bits = _engagement_breakdown(p)
        if metric_bits:
            st.caption(metric_bits)
        url = _safe(p.get("URL"))
        if url:
            st.markdown(f"[View source ↗]({url})")


def _social_metric_chips(row):
    chips = []
    for label, col in [("Likes", "Likes"), ("Comments", "Comments"), ("Shares", "Shares"), ("Views", "Views")]:
        try:
            value = int(row.get(col) or 0)
        except Exception:
            value = 0
        if value:
            chips.append(f"<span class='social-kpi'>{label} {value:,}</span>")
    return "".join(chips)


def render_platform_post_card(p):
    """Premium card for a single owned social item."""
    bucket = theme_bucket(_safe(p.get("Theme")))
    img_html = post_media_html(p, bucket=bucket, source_type=_safe(p.get("Source type"), "Social post"))
    fmt = _safe(p.get("Content format")) or post_format(p.get("URL"))
    sent = _safe(p.get("Sentiment"))
    chips_html = chip(bucket, ATLAS_BLUE) + chip(fmt, "#6B7280")
    if sent:
        chips_html += chip(sent, SENTIMENT_COLORS.get(sent, GRAY))

    title = _safe(p.get("Title")) or _safe(p.get("Summary"))[:100] or "Untitled post"
    date_str = _display_date(p.get("Date"))
    try:
        eng = int(p.get("Engagement") or 0)
    except Exception:
        eng = 0
    score_html = f"<span class='post-card-score'>{eng:,} engagement</span>" if eng else ""
    metrics_html = _social_metric_chips(p)
    metrics_html = f"<div class='post-card-metrics'>{metrics_html}</div>" if metrics_html else ""
    url = _safe(p.get("URL"))
    link_html = f"<div class='source-link'><a href='{escape(url)}' target='_blank'>View source &rarr;</a></div>" if url else ""

    st.markdown(
        f"<div class='post-card'>{img_html}<div class='post-card-body'>"
        f"{chips_html}<div class='post-card-title'>{escape(title[:150])}</div>"
        f"{score_html}<div class='post-card-meta'>{escape(date_str)}</div>{metrics_html}{link_html}"
        f"</div></div>",
        unsafe_allow_html=True,
    )


def render_platform_section(posts_df, key_prefix):
    """Filter row (All + whichever content-type buckets actually have posts)
    plus a 2-column card grid for one platform tab. Shows the soft empty
    message instead of a dead filter row when there's nothing to show."""
    if posts_df.empty:
        st.info("No signals found this period.")
        return
    posts_df = posts_df.copy()
    posts_df["_bucket"] = posts_df["Theme"].apply(theme_bucket)
    posts_df["Engagement"] = pd.to_numeric(posts_df["Engagement"], errors="coerce").fillna(0)
    top_bucket = posts_df["_bucket"].value_counts().index[0] if not posts_df.empty else "No bucket"
    best_post = posts_df.sort_values("Engagement", ascending=False).iloc[0] if not posts_df.empty else None
    s1, s2, s3 = st.columns(3)
    signal_card(s1, "Posts / videos", f"{len(posts_df):,}", "Items in this platform tab.")
    signal_card(s2, "Total engagement", f"{int(posts_df['Engagement'].sum()):,}",
                "Likes, comments, shares, reactions, and views where available.", accent=ATLAS_BLUE_DARK)
    signal_card(s3, "Largest content bucket", top_bucket,
                f"{int((posts_df['_bucket'] == top_bucket).sum())} items." if top_bucket != "No bucket" else "", accent=ACCENT_AMBER)
    if best_post is not None and int(best_post.get("Engagement") or 0):
        insight(
            f"The strongest item in this tab is a <b>{theme_bucket(_safe(best_post.get('Theme')))}</b> post "
            f"with <b>{int(best_post.get('Engagement') or 0):,}</b> tracked engagement. "
            "Use the content-bucket filter below to compare the type of message, not just the channel."
        )
    present = [b for b in THEME_BUCKET_ORDER if (posts_df["_bucket"] == b).any()]
    options = ["All"] + present
    st.markdown("<div class='platform-filter-label'>Content bucket filter</div>", unsafe_allow_html=True)
    choice = st.radio(
        "Content bucket filter", options, index=0, horizontal=True,
        key=f"filter_{key_prefix}", label_visibility="collapsed",
    )
    shown = posts_df if choice == "All" else posts_df[posts_df["_bucket"] == choice]
    shown = shown.sort_values("_dt", ascending=False)
    rows = list(shown.iterrows())
    for i in range(0, len(rows), 2):
        row_slice = rows[i:i + 2]
        cols = st.columns(2)
        for col, (_, p) in zip(cols, row_slice):
            with col:
                render_platform_post_card(p)


def render_post_card(p):
    """Premium card for blog posts, third-party mentions, and other signals."""
    theme = _safe(p.get("Theme"))
    img_html = post_media_html(p, bucket=theme_bucket(theme), source_type=_safe(p.get("Source type"), "Signal"))
    sent = _safe(p.get("Sentiment"))
    source_type = _safe(p.get("Source type"), "Signal")
    chips_html = ""
    if theme:
        chips_html += chip(theme, "#6B7280")
    if sent:
        chips_html += chip(sent, SENTIMENT_COLORS.get(sent, GRAY))
    title = _safe(p.get("Title")) or _safe(p.get("Summary"))[:100] or "Untitled signal"
    date_str = _display_date(p.get("Date"))
    try:
        eng = int(p.get("Engagement") or 0)
    except Exception:
        eng = 0
    score_html = f"<span class='post-card-score'>{eng:,} engagement</span>" if eng else ""
    metrics_html = _social_metric_chips(p)
    metrics_html = f"<div class='post-card-metrics'>{metrics_html}</div>" if metrics_html else ""
    url = _safe(p.get("URL"))
    link_html = f"<div class='source-link'><a href='{escape(url)}' target='_blank'>View source &rarr;</a></div>" if url else ""
    st.markdown(
        f"<div class='post-card'>{img_html}<div class='post-card-body'>"
        f"<div class='source-eyebrow'>{escape(source_type)}</div>{chips_html}"
        f"<div class='post-card-title'>{escape(title[:150])}</div>"
        f"{score_html}<div class='post-card-meta'>{escape(date_str)}</div>{metrics_html}{link_html}"
        f"</div></div>",
        unsafe_allow_html=True,
    )


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


def render_hiring_signal_section(name, h):
    if info_if_empty(h, f"hiring & expansion at {name}"):
        return
    h = h.copy()
    h["_dt"] = pd.to_datetime(h.get("Date"), errors="coerce")
    h = h.sort_values("_dt", ascending=False, na_position="last")

    locations = sorted({_safe(v) for v in h.get("Location (HQ)", pd.Series(dtype=str)) if _safe(v)})
    functions = h.get("Top open-role function (latest snapshot)", pd.Series(dtype=str)).dropna()
    lead_function = _safe(functions.mode().iloc[0]) if not functions.empty else "Not specified"

    c1, c2, c3 = st.columns(3)
    signal_card(c1, "Hiring / expansion signals", f"{len(h):,}",
                "Rows confirmed in the cleaned source data.")
    signal_card(c2, "Locations surfaced", f"{len(locations):,}",
                _join_natural(locations[:3]) if locations else "No location field available.", accent=ACCENT_GREEN)
    signal_card(c3, "Most common role focus", lead_function,
                "Based on the latest open-role function field.", accent=ACCENT_AMBER)

    rows = list(h.head(6).iterrows())
    for i in range(0, len(rows), 3):
        cols = st.columns(3)
        for col, (_, r) in zip(cols, rows[i:i + 3]):
            meta = [
                _safe(r.get("Location (HQ)")),
                _display_date(r.get("Date")),
                _safe(r.get("Team to review")),
            ]
            note = _safe(r.get("Possible meaning")) or _safe(r.get("Signal detail"))
            source_card(
                col,
                _safe(r.get("Signal type"), "Hiring signal"),
                _safe(r.get("Signal detail")) or _safe(r.get("Top open-role function (latest snapshot)"), "Hiring / expansion signal"),
                meta,
                note,
                _safe(r.get("URL")),
                accent=ACCENT_GREEN,
            )
    compact_detail_expander(
        f"View all {len(h)} hiring / expansion rows",
        h,
        ["Signal type", "Location (HQ)", "Top open-role function (latest snapshot)", "Date", "URL",
         "Signal detail", "Possible meaning", "Team to review"],
    )


def render_pr_signal_section(name, pr):
    if info_if_empty(pr, f"PR, news, events & webinars for {name}"):
        return
    pr = pr.copy()
    pr["_dt"] = pd.to_datetime(pr.get("Date"), errors="coerce")
    pr = pr.sort_values("_dt", ascending=False, na_position="last")

    source_counts = pr.get("Source channel", pd.Series(dtype=str)).value_counts()
    type_counts = pr.get("Type", pd.Series(dtype=str)).value_counts()
    top_source = _safe(source_counts.index[0]) if not source_counts.empty else "No source"
    top_type = _safe(type_counts.index[0]) if not type_counts.empty else "No type"
    latest = _display_date(pr["_dt"].max()) if pr["_dt"].notna().any() else "No date"

    c1, c2, c3, c4 = st.columns(4)
    signal_card(c1, "Total coverage rows", f"{len(pr):,}", "PR, news, events, and webinar rows.")
    signal_card(c2, "Top source", top_source, f"{int(source_counts.iloc[0]) if not source_counts.empty else 0} rows.", accent=ATLAS_BLUE_DARK)
    signal_card(c3, "Top signal type", top_type, f"{int(type_counts.iloc[0]) if not type_counts.empty else 0} rows.", accent=ACCENT_AMBER)
    signal_card(c4, "Latest dated item", latest, "Most recent dated coverage row.", accent=ACCENT_GREEN)

    rows = list(pr.head(6).iterrows())
    for i in range(0, len(rows), 3):
        cols = st.columns(3)
        for col, (_, r) in zip(cols, rows[i:i + 3]):
            meta = [
                _safe(r.get("Source channel")),
                _safe(r.get("Type")),
                _display_date(r.get("Date")),
            ]
            chips_html = ""
            theme = _safe(r.get("Theme"))
            if theme:
                chips_html = chip(theme, ATLAS_BLUE_LIGHT, ATLAS_BLUE_DARK)
            source_card(
                col,
                _safe(r.get("Source channel"), "Coverage"),
                _safe(r.get("Title"), "Coverage item"),
                meta,
                _safe(r.get("Possible meaning")),
                _safe(r.get("URL")),
                accent=ATLAS_BLUE,
                chips_html=chips_html,
            )
    compact_detail_expander(
        f"View all {len(pr)} PR / news / event / webinar rows",
        pr,
        ["Source channel", "Type", "Title", "Date", "URL", "Theme", "Possible meaning", "Team to review"],
    )


def render_product_signal_section(name, ps):
    if info_if_empty(ps, f"product & service signals for {name}"):
        return
    ps = ps.copy()
    ps["_dt"] = pd.to_datetime(ps.get("Date found/published"), errors="coerce")
    ps = ps.sort_values("_dt", ascending=False, na_position="last")

    cat_counts = ps.get("Product/service category", pd.Series(dtype=str)).value_counts()
    top_cat = _safe(cat_counts.index[0]) if not cat_counts.empty else "No category"
    latest = _display_date(ps["_dt"].max()) if ps["_dt"].notna().any() else "No date"

    c1, c2, c3 = st.columns(3)
    signal_card(c1, "Product / service rows", f"{len(ps):,}", "Confirmed rows in this month's cleaned data.")
    signal_card(c2, "Most common category", top_cat,
                f"{int(cat_counts.iloc[0]) if not cat_counts.empty else 0} rows.", accent=ACCENT_AMBER)
    signal_card(c3, "Latest dated item", latest, "Most recent product/service row.", accent=ACCENT_GREEN)

    rows = list(ps.head(6).iterrows())
    for i in range(0, len(rows), 3):
        cols = st.columns(3)
        for col, (_, r) in zip(cols, rows[i:i + 3]):
            category = _safe(r.get("Product/service category"), "Product / service")
            chips_html = chip(category, ATLAS_BLUE_LIGHT, ATLAS_BLUE_DARK)
            meta = [_display_date(r.get("Date found/published")), _safe(r.get("Team to review"))]
            source_card(
                col,
                "Product / service signal",
                _safe(r.get("Page/post title"), category),
                meta,
                _safe(r.get("Possible meaning")),
                _safe(r.get("URL")),
                accent=ACCENT_AMBER,
                chips_html=chips_html,
            )
    compact_detail_expander(
        f"View all {len(ps)} product / service rows",
        ps,
        ["Product/service category", "Page/post title", "URL", "Date found/published", "Possible meaning", "Team to review"],
    )


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

    section_divider()
    brief_header(
        "Content performance read",
        f"Four connected views of {name}'s own posts only: engagement trend, engagement by theme, posting mix by theme, and sentiment mix.",
        "Owned content analysis",
    )
    if False:
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
    insight(
        "Read the four charts as a sequence: when engagement happened, which themes performed, what they posted most, "
        "and whether the tone was positive, neutral, mixed, or negative. This keeps the posting mix and sentiment mix tied "
        "to the performance story instead of floating as standalone charts."
    )
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
    brief_row = _june_brief_row(name)

    level = _safe(row.get("Overall activity level"), "n/a")
    level_color = ACTIVITY_LEVEL_COLORS.get(level, GRAY)
    total_signals = _num(brief_row, "Total June signals", int(row["Total activity"]))
    social_posts_n = _num(brief_row, "Social posts")
    social_eng_n = _num(brief_row, "Social engagement")
    avg_social_n = brief_row.get("Avg social engagement/post", 0) if brief_row is not None else 0

    hcol1, hcol2 = st.columns([1, 7])
    with hcol1:
        st.markdown(logo_badge(name, LOGO_MAP, 72), unsafe_allow_html=True)
    with hcol2:
        st.markdown(f"### {name}")
        st.markdown(chip(level + " activity", level_color), unsafe_allow_html=True)
        st.caption(
            f"{total_signals} June signals tracked across owned content, social activity, news, and workforce signals."
        )

    st.write("")
    k1, k2, k3, k4 = st.columns(4)
    kpi_card(k1, "\U0001F4CA", "June signals", total_signals)
    kpi_card(k2, "\U0001F525", "Social engagement", f"{social_eng_n:,}",
             f"{float(avg_social_n):.1f} avg / post" if social_posts_n else None)
    kpi_card(k3, "\U0001F4DD", "Blogs + news", _num(brief_row, "Blog posts") + _num(brief_row, "News items"))
    kpi_card(k4, "\U0001F3A5", "Training/webinar signals", _num(brief_row, "Training/webinar-themed signals"))

    p1, p2, p3, p4 = st.columns(4)
    mini_stat(p1, "LinkedIn", f"{_num(brief_row, 'LinkedIn posts', int(row.get('LinkedIn posts', 0) or 0))} posts")
    mini_stat(p2, "Instagram", f"{_num(brief_row, 'Instagram posts', int(row.get('Instagram posts', 0) or 0))} posts")
    mini_stat(p3, "YouTube", f"{_num(brief_row, 'YouTube videos', int(row.get('YouTube videos', 0) or 0))} videos")
    mini_stat(p4, "High-priority signals", f"{_num(brief_row, 'High-priority signals')} for review")
    profile_source = june_content_items.copy() if not june_content_items.empty else raw_data.copy()
    if "Month" in profile_source.columns:
        profile_source = profile_source[profile_source["Month"] == settings.get("Report_Month", "June 2026")]

    section_divider()
    brief_header(
        "Channel footprint and themes",
        f"Where {name} showed up in June, plus the strongest themes attached to those signals.",
        "Competitor activity map",
    )
    cc1, cc2 = st.columns([3, 2])
    with cc1:
        chans = {
            "LinkedIn": _num(brief_row, "LinkedIn posts", int(row.get("LinkedIn posts", 0) or 0)),
            "Instagram": _num(brief_row, "Instagram posts", int(row.get("Instagram posts", 0) or 0)),
            "YouTube": _num(brief_row, "YouTube videos", int(row.get("YouTube videos", 0) or 0)),
            "Blog": _num(brief_row, "Blog posts"),
            "News": _num(brief_row, "News items"),
        }
        cdf = pd.DataFrame({"Channel": list(chans.keys()), "Count": list(chans.values())})
        cdf = cdf[cdf["Count"] > 0]
        if cdf.empty:
            st.info("No signals found this period for LinkedIn, Instagram, or YouTube activity.")
        else:
            cdf = highlight_leader(cdf.sort_values("Count"), "Count")
            fig = px.bar(cdf, x="Count", y="Channel", orientation="h", text="Count", color="_Highlight",
                         color_discrete_map={"Leading": ATLAS_BLUE, "Other": NEUTRAL_BAR},
                         title="June channel activity")
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
    if not profile_source.empty:
        _comp_posts = profile_source[(profile_source["Competitor"] == name) & (profile_source["Channel"] != "Jobs")]
        owned_for_chart = _comp_posts[_comp_posts["Source type"].isin(OWNED_SOURCE_TYPES)]
    render_engagement_section(name, owned_for_chart)

    section_divider()
    brief_header(
        "Hiring and expansion signals",
        "Workforce and footprint clues that may indicate where the competitor is investing attention.",
        "Operational signals",
    )
    h = hiring[hiring["Competitor"] == name] if not hiring.empty else pd.DataFrame()
    render_hiring_signal_section(name, h)

    section_divider()
    brief_header(
        "PR, news, events and webinars",
        "Earned or scheduled visibility for the month, presented as source cards with the full table available on demand.",
        "Market visibility",
    )
    pr = pr_news[pr_news["Competitor"] == name] if not pr_news.empty else pd.DataFrame()
    if not pr.empty and "Month" in pr.columns:
        pr = pr[pr["Month"] == settings.get("Report_Month", "June 2026")]
    render_pr_signal_section(name, pr)

    section_divider()
    brief_header(
        "Product and service signals",
        "Pages, posts, and content items that point to product, service, aftermarket, or category emphasis.",
        "Offer signals",
    )
    ps = product_signals[product_signals["Competitor"] == name] if not product_signals.empty else pd.DataFrame()
    if not ps.empty and "Month" in ps.columns:
        ps = ps[ps["Month"] == settings.get("Report_Month", "June 2026")]
    render_product_signal_section(name, ps)

    section_divider()
    brief_header(
        "Opportunities for review",
        "Competitor-specific opportunities are shown here when the cleaned opportunity sheet explicitly names this brand.",
        "Review focus",
    )
    matches = pd.DataFrame()
    if not opportunities.empty:
        matches = opportunities[opportunities["Opportunity/Signal"].str.contains(name, case=False, na=False)]
    if matches.empty:
        st.caption(f"No opportunity specifically flagged for {name} this period — see the Executive Summary "
                   "for cross-competitor signals.")
    else:
        for _, r in matches.iterrows():
            p_color = PRIORITY_COLORS.get(r["Priority"], GRAY)
            with st.container(border=True):
                st.markdown(f"**{r['Opportunity/Signal']}**")
                st.markdown(
                    f"<div class='card-meta'><span class='meta-dot' style='background:{p_color};'></span>"
                    f"{r['Priority']} priority &middot; {r['Confidence']} confidence &middot; "
                    f"Suggested team: {r['Suggested team to review']}</div>",
                    unsafe_allow_html=True,
                )
                st.markdown(f"*Why it may matter:* {r['Why it may matter']}")
                with st.expander("Evidence"):
                    st.write(r["Evidence"])

    st.markdown(
        "##### Social media activity"
        "<span class='section-hint'>· their own channels, plus third-party mentions</span>",
        unsafe_allow_html=True,
    )
    all_posts = profile_source[profile_source["Competitor"] == name].copy() if not profile_source.empty else pd.DataFrame()
    all_posts = all_posts[all_posts["Channel"] != "Jobs"] if not all_posts.empty else all_posts
    if all_posts.empty:
        st.info("No signals found this period.")
    else:
        all_posts["_dt"] = pd.to_datetime(all_posts["Date"], errors="coerce")
        all_posts["Engagement"] = pd.to_numeric(all_posts["Engagement"], errors="coerce").fillna(0)
        owned_total = int(all_posts["Source type"].isin(OWNED_SOURCE_TYPES).sum())
        mention_total = int(all_posts["Source type"].isin(MENTION_SOURCE_TYPES).sum())
        source_counts = all_posts["Source type"].value_counts()
        top_source = _safe(source_counts.index[0]) if not source_counts.empty else "No source"
        s1, s2, s3, s4 = st.columns(4)
        signal_card(s1, "Owned posts / videos", f"{owned_total:,}", "Company or employee channel items.")
        signal_card(s2, "Third-party mentions", f"{mention_total:,}", "Posts and comments mentioning the brand.", accent=ACCENT_AMBER)
        signal_card(s3, "Tracked engagement", f"{int(all_posts['Engagement'].sum()):,}",
                    "Engagement fields available in the source rows.", accent=ATLAS_BLUE_DARK)
        signal_card(s4, "Top source type", top_source,
                    f"{int(source_counts.iloc[0]) if not source_counts.empty else 0} rows.", accent=ACCENT_GREEN)

        # Coverage strip: one glance at what showed up where before opening anything -
        # lit up in blue when there's content, muted gray when there's none.
        strip_html = ""
        for label, types, icon in OWNED_CHANNEL_BUCKETS + MENTION_TYPE_BUCKETS:
            n = int(all_posts["Source type"].isin(types).sum())
            strip_html += chip(f"{icon} {label} · {n}", ATLAS_BLUE if n else NEUTRAL_BAR, fg=WHITE if n else NAVY)
        st.markdown(strip_html, unsafe_allow_html=True)
        st.write("")

        platform_tabs = st.tabs([
            "\U0001F4BC LinkedIn", "\U0001F4F7 Instagram", "▶️ YouTube", "\U0001F9E9 Other & mentions",
        ])

        for (platform_label, types), tab in zip(PLATFORM_SOURCE_TYPES, platform_tabs[:3]):
            with tab:
                sub = all_posts[all_posts["Source type"].isin(types)]
                render_platform_section(sub, key_prefix=f"{name}_{platform_label}")

        with platform_tabs[3]:
            blog_posts = all_posts[all_posts["Source type"].isin({"Blog pages/articles"})]
            if not blog_posts.empty:
                st.markdown("**\U0001F4DD Blog**")
                render_post_section(blog_posts)
                st.write("")

            mention_total = int(all_posts["Source type"].isin(MENTION_SOURCE_TYPES).sum())
            st.caption(f"Other people's posts and comments using {name}'s name - {mention_total} found this period.")
            any_mention = False
            for mention_label, types, icon in MENTION_TYPE_BUCKETS:
                sub = all_posts[all_posts["Source type"].isin(types)]
                if sub.empty:
                    continue
                any_mention = True
                st.markdown(f"**{icon} {mention_label}**")
                render_post_section(sub)
                st.write("")
            if not any_mention and blog_posts.empty:
                st.caption(f"No blog posts or third-party mentions of {name} were found this period.")

            covered_types = OWNED_SOURCE_TYPES | MENTION_SOURCE_TYPES
            leftover = all_posts[~all_posts["Source type"].isin(covered_types)]
            if not leftover.empty:
                st.markdown("**\U0001F9E9 Other signals**")
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
        br = _june_brief_row(nm)
        rows.append({
            "Logo": LOGO_MAP.get(nm),
            "Competitor": nm,
            "Activity level": _safe(r.get("Overall activity level"), "n/a"),
            "June signals": _num(br, "Total June signals", int(r["Total activity"])),
            "LinkedIn": _num(br, "LinkedIn posts", int(r.get("LinkedIn posts", 0) or 0)),
            "Instagram": _num(br, "Instagram posts", int(r.get("Instagram posts", 0) or 0)),
            "YouTube": _num(br, "YouTube videos", int(r.get("YouTube videos", 0) or 0)),
            "News": _num(br, "News items"),
            "Social engagement": _num(br, "Social engagement"),
            "Top channel": _top_channel_for(r),
            "Top theme": _safe(br.get("Top non-general theme")) if br is not None else _top_theme_for(nm),
        })
    odf = pd.DataFrame(rows).sort_values("June signals", ascending=False)

    if not odf.empty:
        leader = odf.iloc[0]
        insight(
            f"Among the remaining tracked competitors, <b>{leader['Competitor']}</b> is the most active with "
            f"{leader['June signals']} June signals this period — still below the top competitors shown in their own tabs."
        )

    st.dataframe(
        odf, width="stretch", hide_index=True,
        column_config={
            "Logo": st.column_config.ImageColumn("Logo"),
        },
    )
    st.caption("Logos shown only where a source image was available in the scraped data — no logo is shown "
               "rather than guessing.")

    st.markdown("##### PR & news mentions")
    pr_other = pr_news[pr_news["Competitor"].isin(OTHER_COMPETITORS)] if not pr_news.empty else pd.DataFrame()
    if not pr_other.empty and "Month" in pr_other.columns:
        pr_other = pr_other[pr_other["Month"] == settings.get("Report_Month", "June 2026")]
    if info_if_empty(pr_other, "PR, news, or event mentions for these competitors"):
        pass
    else:
        st.dataframe(
            pr_other[["Competitor", "Source channel", "Type", "Title", "Date", "URL", "Theme"]]
            .sort_values("Date").reset_index(drop=True),
            width="stretch", hide_index=True,
            column_config={"URL": st.column_config.LinkColumn("URL")},
        )
        _covered_other = sorted(pr_other["Competitor"].unique().tolist())
        _no_cov_other = [c for c in OTHER_COMPETITORS if c not in _covered_other]
        if _no_cov_other:
            st.caption("No PR, news, or event mentions found this period for: " + ", ".join(_no_cov_other) + ".")


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

    # ---- June intelligence cockpit: richer sheets built by the ETL --------
    if not june_brief.empty:
        section_divider()
        brief_header(
            "June competitor snapshot",
            "A ranked view of signal volume, social traction, content output, news visibility, and priority flags.",
            "Executive intelligence cockpit",
        )
        report_month = settings.get("Report_Month", "June 2026")
        jb = june_brief.sort_values("Total June signals", ascending=False).reset_index(drop=True)
        total_june = int(jb["Total June signals"].sum())
        total_social_eng = int(jb["Social engagement"].sum())
        leader = jb.iloc[0]
        social_leader = jb.sort_values("Social engagement", ascending=False).iloc[0]

        c1, c2, c3, c4 = st.columns(4)
        kpi_card(c1, "\U0001F9ED", f"{report_month} signals", f"{total_june:,}",
                 f"{len(jb)} tracked competitors")
        kpi_card(c2, "\U0001F525", "Social engagement", f"{total_social_eng:,}",
                 f"{social_leader['Competitor']} led engagement")
        kpi_card(c3, "\U0001F4DD", "Blogs + news",
                 f"{int(jb['Blog posts'].sum() + jb['News items'].sum()):,}",
                 "confirmed dated items only")
        kpi_card(c4, "\U0001F3A5", "Webinar/training signals",
                 f"{int(jb['Training/webinar-themed signals'].sum()):,}",
                 f"{int(jb['Confirmed webinar listings'].sum())} confirmed webinar listings")

        insight(
            f"<b>{leader['Competitor']}</b> had the highest June signal volume ({int(leader['Total June signals'])}), "
            f"while <b>{social_leader['Competitor']}</b> led total social engagement "
            f"({int(social_leader['Social engagement']):,}). Confirmed webinar listings and confirmed PR releases "
            f"are kept separate from broader training/news signals to avoid overstating the evidence."
        )

        pcol1, pcol2 = st.columns([3, 2])
        with pcol1:
            show_cols = [
                "Competitor", "Total June signals", "Social posts", "Social engagement",
                "Blog posts", "News items", "Training/webinar-themed signals",
                "High-priority signals", "Top non-general theme",
            ]
            snap = jb[[c for c in show_cols if c in jb.columns]].head(9).copy()
            max_signals = max(int(snap["Total June signals"].max() or 1), 1)
            max_engagement = max(int(snap["Social engagement"].max() or 1), 1)
            row_html = [
                "<div class='snapshot-table'>"
                "<div class='snapshot-row snapshot-head'><div>Rank</div><div>Competitor</div>"
                "<div>June signals</div><div>Social engagement</div><div>Lead theme</div></div>"
            ]
            for rank, (_, r) in enumerate(snap.iterrows(), start=1):
                signals = int(r.get("Total June signals") or 0)
                engagement = int(r.get("Social engagement") or 0)
                sig_width = signals / max_signals * 100
                eng_width = engagement / max_engagement * 100
                row_html.append(
                    "<div class='snapshot-row'>"
                    f"<div class='snapshot-rank'>{rank}</div>"
                    f"<div><div class='snapshot-name'>{escape(_safe(r.get('Competitor')))}</div>"
                    f"<div class='snapshot-small'>{int(r.get('Social posts') or 0)} social posts · "
                    f"{int(r.get('High-priority signals') or 0)} high priority</div></div>"
                    f"<div><div class='snapshot-small'>{signals:,}</div>{simple_bar(signals, max_signals)}</div>"
                    f"<div><div class='snapshot-small'>{engagement:,}</div>{simple_bar(engagement, max_engagement, ATLAS_BLUE_DARK)}</div>"
                    f"<div class='snapshot-small'>{escape(_safe(r.get('Top non-general theme'), 'General brand activity'))}</div>"
                    "</div>"
                )
            row_html.append("</div>")
            st.markdown("".join(row_html), unsafe_allow_html=True)
        with pcol2:
            if not june_content_mix.empty:
                mix = june_content_mix.copy()
                mix["Owned content"] = mix[["Social media posts", "Blog posts", "Product/service signals"]].sum(axis=1)
                fig = px.bar(
                    mix.sort_values("Owned content", ascending=True),
                    x="Owned content", y="Competitor", orientation="h", text="Owned content",
                    color_discrete_sequence=[ATLAS_BLUE],
                    title="Owned/content signal volume",
                )
                fig.update_yaxes(title="")
                fig.update_traces(textposition="outside", cliponaxis=False)
                clean_value_axis(fig, "h")
                st.plotly_chart(style_layout(fig, 360, legend="none"), width="stretch")

    if not june_social_metrics.empty:
        section_divider()
        brief_header(
            "Social performance by platform",
            "Engagement and posting cadence by competitor-platform combination.",
            "Channel performance",
        )
        sm = june_social_metrics.copy()
        sm["Engagement"] = pd.to_numeric(sm["Engagement"], errors="coerce").fillna(0)
        sm["Posts/videos"] = pd.to_numeric(sm["Posts/videos"], errors="coerce").fillna(0)
        scol1, scol2 = st.columns(2)
        with scol1:
            fig = px.bar(
                sm.sort_values("Engagement", ascending=True).tail(12),
                x="Engagement", y="Competitor", color="Platform", orientation="h",
                text="Engagement", color_discrete_sequence=CATEGORICAL_PALETTE,
                title="Highest social engagement rows",
            )
            fig.update_yaxes(title="")
            fig.update_traces(textposition="outside", cliponaxis=False)
            clean_value_axis(fig, "h")
            st.plotly_chart(style_layout(fig, 390, legend="bottom"), width="stretch")
        with scol2:
            fig = px.scatter(
                sm, x="Posts/videos", y="Engagement", size="Avg engagement/post",
                color="Platform", hover_name="Competitor",
                color_discrete_sequence=CATEGORICAL_PALETTE,
                title="Cadence vs engagement",
            )
            fig.update_xaxes(title="Posts / videos")
            fig.update_yaxes(title="Engagement")
            st.plotly_chart(style_layout(fig, 390, legend="bottom"), width="stretch")

        top_social = sm.sort_values("Engagement", ascending=False).head(6)
        brief_header(
            "Top social performance cards",
            "The highest-engagement competitor-platform rows, with the strongest post or video surfaced for review.",
            "Social detail",
            compact=True,
        )
        rows = list(top_social.iterrows())
        for i in range(0, len(rows), 3):
            cols = st.columns(3)
            for col, (_, r) in zip(cols, rows[i:i + 3]):
                title = _safe(r.get("Top post/video"), "No title available")
                url = _safe(r.get("Top post/video URL"))
                link_html = f"<div class='news-meta'><a href='{url}' target='_blank'>View source ↗</a></div>" if url else ""
                metric_html = "".join([
                    f"<span class='social-kpi'>Posts {int(r.get('Posts/videos') or 0):,}</span>",
                    f"<span class='social-kpi'>Engagement {int(r.get('Engagement') or 0):,}</span>",
                    f"<span class='social-kpi'>Likes {int(r.get('Likes') or 0):,}</span>",
                    f"<span class='social-kpi'>Comments {int(r.get('Comments') or 0):,}</span>",
                    f"<span class='social-kpi'>Shares {int(r.get('Shares') or 0):,}</span>" if int(r.get("Shares") or 0) else "",
                    f"<span class='social-kpi'>Views {int(r.get('Views') or 0):,}</span>" if int(r.get("Views") or 0) else "",
                ])
                col.markdown(
                    f"<div class='social-card'>"
                    f"<div class='social-card-meta'>{r['Competitor']} · {r['Platform']}</div>"
                    f"<div class='social-card-title'>{title}</div>"
                    f"<div class='social-card-kpis'>{metric_html}</div>{link_html}</div>",
                    unsafe_allow_html=True,
                )

    # ---- The Big Picture: 4 synthesis charts -------------------------------
    section_divider()
    brief_header(
        "The big picture",
        "Four board-level views that summarize activity volume, channel mix, messaging themes, and market tone.",
        "Strategic read",
    )
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

    # ---- Signals snapshot: calmer scorecards -------------------------------
    section_divider()
    brief_header(
        "Other signals worth knowing",
        "Secondary indicators that may be useful for context, but should be validated before action.",
        "Signal context",
    )
    s1, s2, s3, s4 = st.columns(4)
    if not brand_mentions.empty:
        bm = brand_mentions[~brand_mentions["Brand"].isin(["Unspecified / General"])].sort_values("Mentions", ascending=False)
        if not bm.empty:
            top_bm = bm.iloc[0]
            signal_card(
                s1, "Most mentioned brand", top_bm["Brand"],
                f"{int(top_bm['Mentions'])} mentions in social conversation.",
            )
    if not hiring.empty:
        hc = hiring["Competitor"].value_counts()
        signal_card(
            s2, "Hiring / workforce", hc.index[0],
            f"{int(hc.iloc[0])} of {len(hiring)} hiring or expansion signals.",
            accent=ACCENT_GREEN,
        )
    confirmed_webinars = int(june_brief["Confirmed webinar listings"].sum()) if not june_brief.empty else 0
    confirmed_pr = int(june_brief["Confirmed PR releases"].sum()) if not june_brief.empty else 0
    coverage_note = (
        "No confirmed webinar listings or PR releases found in the cleaned June data."
        if confirmed_webinars == 0 and confirmed_pr == 0
        else f"{confirmed_webinars} confirmed webinar listings and {confirmed_pr} confirmed PR releases."
    )
    signal_card(s3, "Coverage to validate", f"{confirmed_webinars + confirmed_pr}", coverage_note, accent=ACCENT_AMBER)
    our_brand_name = str(settings.get("Our_Brand", "Atlas Copco"))
    if not brand_mentions.empty and our_brand_name in brand_mentions["Brand"].values:
        our_mentions = int(brand_mentions.loc[brand_mentions["Brand"] == our_brand_name, "Mentions"].iloc[0])
        signal_card(
            s4, "Our brand mentions", f"{our_mentions}",
            f"{our_brand_name} appeared in the tracked conversation.",
            accent=ATLAS_BLUE_DARK,
        )

    # ---- Key themes to watch, as ranked briefing cards ---------------------
    section_divider()
    brief_header(
        "Key themes to watch",
        "The strongest non-general themes found in competitor content, ranked by frequency.",
        "Messaging themes",
    )
    theme_lookup = {}
    total_themed = 0
    if not theme_freq.empty:
        total_themed = theme_freq.loc[theme_freq["Theme"] != "General brand activity", "Frequency"].sum()
        theme_lookup = dict(zip(theme_freq["Theme"], theme_freq["Frequency"]))
    theme_rows = []
    for i in (1, 2, 3):
        theme_name = exec_metric(f"Top theme #{i}")
        theme_rows.append((i, theme_name, theme_lookup.get(theme_name)))
    top_freq = max((f for _, _, f in theme_rows if f), default=None)
    tcols = st.columns(3)
    for (i, theme_name, freq), col in zip(theme_rows, tcols):
        share = _pct(freq, total_themed) if freq and total_themed else "—"
        theme_note = ""
        if not raw_data.empty and "Theme" in raw_data.columns and "Possible meaning" in raw_data.columns:
            note_rows = raw_data.loc[raw_data["Theme"] == theme_name, "Possible meaning"].dropna()
            if not note_rows.empty:
                theme_note = _safe(note_rows.iloc[0])
        with col:
            st.markdown(
                f"<div class='rank-card'>"
                f"<div class='rank-top'><div class='rank-num'>{i}</div><div class='rank-title'>{theme_name}</div></div>"
                f"<div class='rank-metric'>{share}</div>"
                f"{simple_bar(freq or 0, top_freq or 0)}"
                f"<div class='rank-copy' style='margin-top:8px;'>{int(freq or 0)} signals. {theme_note}</div>"
                f"</div>",
                unsafe_allow_html=True,
            )

    # ---- Top opportunities, as ranked review cards -------------------------
    section_divider()
    brief_header(
        "Top opportunities for review",
        "Data-backed signals that may be worth a cross-functional look.",
        "Review focus",
    )
    ocols = st.columns(3)
    for i, col in zip((1, 2, 3), ocols):
        opp_text = exec_metric(f"Possible opportunity #{i}")
        match = pd.DataFrame()
        if not opportunities.empty and opp_text and opp_text != "n/a":
            match = opportunities[opportunities["Opportunity/Signal"] == opp_text]
        if not match.empty:
            r = match.iloc[0]
            accent = PRIORITY_COLORS.get(r["Priority"], ATLAS_BLUE)
            why = _safe(r.get("Why it may matter"))
            team = _safe(r.get("Suggested team to review"))
            meta = (
                f"<div class='card-meta'><span class='meta-dot' style='background:{accent};'></span>"
                f"{r['Priority']} priority &middot; {r['Confidence']} confidence</div>"
                f"<div class='rank-copy'>{why}</div>"
                f"<div class='card-meta'><span class='team-tag'>{team}</span></div>"
            )
        else:
            accent, meta = ATLAS_BLUE, ""
        with col:
            st.markdown(
                f"<div class='rank-card' style='border-top:4px solid {accent};'>"
                f"<div class='rank-top'><div class='rank-num'>{i}</div><div class='rank-title'>{opp_text}</div></div>"
                f"{meta}"
                f"</div>",
                unsafe_allow_html=True,
            )

    # ---- What to do next: compact review queue -----------------------------
    section_divider()
    brief_header(
        "What to do next",
        "A short review queue for the relevant teams. These are signals to validate, not final recommendations.",
        "Review queue",
    )
    if not opportunities.empty:
        prio_rank = {"High": 0, "Medium": 1, "Low": 2}
        top_opps = opportunities.copy()
        top_opps["_r"] = top_opps["Priority"].map(prio_rank).fillna(3)
        top_opps = top_opps.sort_values("_r").head(5).reset_index(drop=True)
        qcols = st.columns(2)
        for i, row in top_opps.iterrows():
            teams = [t.strip() for t in str(row["Suggested team to review"]).split(";") if t.strip()]
            team_html = "".join(f"<span class='team-tag'>{t}</span>" for t in teams)
            p_color = PRIORITY_COLORS.get(row["Priority"], ATLAS_BLUE)
            meta_html = (f"<span class='meta-dot' style='background:{p_color};'></span>"
                         f"<span>{row['Priority']} priority</span>{team_html}")
            with qcols[i % 2]:
                st.markdown(
                    f"<div class='action-card' style='border-left-color:{p_color};'>"
                    f"<div class='action-num'>{i + 1}</div>"
                    f"<div><div class='action-title'>{row['Opportunity/Signal']}</div>"
                    f"<div class='action-why'>{row['Why it may matter']}</div>"
                    f"<div class='card-meta'>{meta_html}</div>"
                    f"</div></div>",
                    unsafe_allow_html=True,
                )
    st.caption("Each top competitor's own tab has the full opportunity detail and evidence behind items related to them.")

    # ---- Industry news pulse (Google News coverage, computed live) ---------
    section_divider()
    brief_header(
        "Industry news pulse",
        "Google News coverage separated into general industry, Atlas Copco, and competitor-specific mentions.",
        "Earned visibility",
    )
    _news_all = pr_news[pr_news["Source channel"] == "News"] if not pr_news.empty else pd.DataFrame()
    if info_if_empty(_news_all, "Google News coverage of the air-compressor category"):
        pass
    else:
        _news_general = _news_all[_news_all["Competitor"] == "Unspecified / General"]
        _news_ac = _news_all[_news_all["Competitor"] == "Atlas Copco"]
        _news_comp = _news_all[~_news_all["Competitor"].isin(["Unspecified / General", "Atlas Copco"])]
        _comp_counts = _news_comp["Competitor"].value_counts()
        _no_cov = [c for c in (TOP_COMPETITORS + OTHER_COMPETITORS) if c not in _comp_counts.index]
        SUBSTANTIVE_TYPES = {"Market expansion", "Product promotion"}

        bits = []
        _volume_leader = "No competitor"
        _vl_n = 0
        _vl_share = 0
        _ac_n = len(_news_ac)
        _ac_share = (_news_ac["Type"].isin(SUBSTANTIVE_TYPES).sum() / _ac_n * 100) if _ac_n else None
        if not _comp_counts.empty:
            _volume_leader = _comp_counts.index[0]
            _vl_df = _news_comp[_news_comp["Competitor"] == _volume_leader]
            _vl_n = len(_vl_df)
            _vl_share = (_vl_df["Type"].isin(SUBSTANTIVE_TYPES).sum() / _vl_n * 100) if _vl_n else 0
            if _ac_share is not None and _vl_share < _ac_share:
                bits.append(
                    f"<b>{_volume_leader}</b> generated the most competitor-specific News volume this period "
                    f"({_vl_n} articles), but only <b>{_vl_share:.0f}%</b> of it was a genuine product launch or "
                    f"business move — the rest reads as routine market or stock-analyst commentary. Atlas "
                    f"Copco's own News coverage was smaller in volume ({_ac_n} articles) but <b>{_ac_share:.0f}%</b> "
                    f"substantive (acquisitions, a product launch), so raw article counts alone overstate how much "
                    f"{_volume_leader} actually moved this period."
                )
            else:
                _share_txt = (f", {_vl_share:.0f}% of it a genuine launch or business move"
                              if _vl_n else "")
                bits.append(
                    f"<b>{_volume_leader}</b> led competitor-specific News volume this period with {_vl_n} "
                    f"articles{_share_txt}."
                )
        else:
            bits.append(
                f"All {len(_news_all)} Google News articles this period were general industry/market coverage or "
                f"about Atlas Copco — none named another tracked competitor."
            )
        if _no_cov:
            bits.append(
                f"No Google News coverage at all surfaced this period for: {', '.join(_no_cov)} — worth a "
                f"sanity check with the PR team, since this could be a real visibility gap or just outside what a "
                f"keyword-based news scrape catches."
            )
        insight(" ".join(bits))

        n1, n2, n3, n4 = st.columns(4)
        signal_card(n1, "Total articles", f"{len(_news_all)}", f"{len(_news_general)} general industry items.")
        signal_card(n2, "Competitor leader", _volume_leader, f"{_vl_n} articles; {_vl_share:.0f}% substantive.")
        signal_card(n3, "Atlas Copco mentions", f"{len(_news_ac)}",
                    f"{_ac_share:.0f}% substantive." if _ac_share is not None else "No Atlas Copco news rows found.",
                    accent=ATLAS_BLUE_DARK)
        signal_card(n4, "No competitor coverage", f"{len(_no_cov)}",
                    "Tracked competitors with no confirmed Google News item in this period.",
                    accent=ACCENT_AMBER)

        ncol1, ncol2 = st.columns([2, 3])
        with ncol1:
            if not _comp_counts.empty:
                cdf = _comp_counts.reset_index()
                cdf.columns = ["Competitor", "Articles"]
                cdf = highlight_leader(cdf.sort_values("Articles"), "Articles")
                fig = px.bar(
                    cdf, x="Articles", y="Competitor", orientation="h", text="Articles",
                    color="_Highlight", color_discrete_map={"Leading": ATLAS_BLUE, "Other": NEUTRAL_BAR},
                    title="Competitor-specific news coverage",
                )
                fig.update_yaxes(title="")
                fig.update_traces(textposition="outside", cliponaxis=False)
                clean_value_axis(fig, "h")
                st.plotly_chart(style_layout(fig, 310, legend="none"), width="stretch")
            else:
                st.info("No competitor-specific news coverage found this period.")
        with ncol2:
            featured_news = _news_all.sort_values("Date", ascending=False).head(4)
            rows = list(featured_news.iterrows())
            for i in range(0, len(rows), 2):
                cols = st.columns(2)
                for col, (_, r) in zip(cols, rows[i:i + 2]):
                    url = _safe(r.get("URL"))
                    link_html = f"<div class='news-meta'><a href='{url}' target='_blank'>View source ↗</a></div>" if url else ""
                    col.markdown(
                        f"<div class='news-card'><div class='news-title'>{_safe(r.get('Title'))}</div>"
                        f"<div class='news-meta'>{_safe(r.get('Competitor'))} · {_safe(r.get('Type'))} · {_safe(r.get('Date'))}</div>"
                        f"{link_html}</div>",
                        unsafe_allow_html=True,
                    )

        with st.expander(f"See all {len(_news_all)} Google News articles tracked this period"):
            st.dataframe(
                _news_all[["Competitor", "Type", "Title", "Date", "URL"]].sort_values("Date").reset_index(drop=True),
                width="stretch", hide_index=True,
                column_config={"URL": st.column_config.LinkColumn("URL")},
            )
        _recap = f"{len(_news_all)} articles total this period — {len(_news_general)} general industry/market coverage, {len(_news_ac)} mentioned Atlas Copco, {len(_news_comp)} named a tracked competitor"
        if not _comp_counts.empty:
            _recap += " (" + "; ".join(f"{c} {n}" for c, n in _comp_counts.items()) + ")"
        st.caption(_recap + ".")
        st.caption("Some keyword-bucket matches (e.g. homonym/local-interest results for \"Gardner Denver\" and "
                   "\"Quincy Compressor\" search terms) were manually reviewed and excluded as false positives. "
                   "See the Settings sheet for full methodology notes.")

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
