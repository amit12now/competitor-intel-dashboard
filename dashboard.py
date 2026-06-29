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
final business decisions. Every table includes soft, review-oriented language
("may be worth reviewing", "could indicate", "possible opportunity") rather than
directives. Categories with zero confirmed signals are shown honestly as
"No signals found this period" rather than estimated or invented.

This version is written for a NON-TECHNICAL audience: filters are kept only on
the Raw Data Explorer (tab 10); every other tab tells a plain-English story with
"What this means" / "Worth considering" callouts computed live from the data.
Chart styling favors a single highlighted "leader" bar over busy multi-color
legends, hides redundant value axes (the number is already printed on the bar),
and keeps legends out of the way of titles - so nothing overlaps or looks scattered.
"""
from pathlib import Path

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

DISCLAIMER = (
    "This dashboard surfaces external competitor signals for review. Final actions "
    "should be validated by the relevant marketing, product, sales, or leadership teams."
)

DEFAULT_PATH = Path(__file__).parent / "Competitor_Intelligence_Master.xlsx"

CLEAN_SHEETS = [
    "Clean_Executive_Summary", "Clean_Competitor_Activity", "Clean_Channel_Activity",
    "Clean_Competitor_x_Channel", "Clean_Trend_May_vs_June", "Clean_Theme_Frequency",
    "Clean_Theme_by_Competitor", "Clean_Theme_by_Channel", "Clean_Theme_Trend_May_vs_June",
    "Clean_Product_Service_Signals", "Clean_Social_Conversation", "Clean_Conversation_Top_Themes",
    "Clean_Conversation_Sentiment", "Clean_Most_Mentioned_Brands", "Clean_PR_News_Events",
    "Clean_Hiring_Expansion", "Clean_Opportunities", "Clean_Raw_Data",
]

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
.disclaimer-banner {{
    background-color: {ATLAS_BLUE_LIGHT};
    border-left: 5px solid {ATLAS_BLUE};
    color: {NAVY};
    padding: 10px 16px;
    border-radius: 6px;
    font-size: 0.88rem;
    margin-bottom: 1rem;
}}
.theme-pill {{
    display:inline-block; background:{ATLAS_BLUE_LIGHT}; color:{ATLAS_BLUE_DARK};
    border-radius: 14px; padding: 3px 12px; margin: 2px; font-size:0.82rem; font-weight:600;
}}
.stTabs [data-baseweb="tab-list"] {{ gap: 22px; border-bottom: 2px solid #E2E8F0; }}
.stTabs [data-baseweb="tab"] {{
    background-color: transparent; border-radius: 0; padding: 10px 2px;
    color: {GRAY}; font-weight: 600;
}}
.stTabs [data-baseweb="tab"]:hover {{ color: {ATLAS_BLUE_DARK}; background-color: transparent; }}
.stTabs [aria-selected="true"] {{
    background-color: transparent !important; color: {ATLAS_BLUE_DARK} !important;
    border-bottom: 3px solid {ATLAS_BLUE} !important; font-weight: 700;
}}
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
.team-tag {{
    display:inline-block; background:{NAVY}; color:{WHITE};
    border-radius: 10px; padding: 1px 10px; margin-right: 6px; font-size:0.74rem; font-weight:700;
}}
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
# Small text helpers (for plain-English, data-grounded storytelling)
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
        f"""<div class='kpi-card'>
            <div class='kpi-icon'>{icon}</div>
            <div class='kpi-label'>{label}</div>
            <div class='kpi-value'>{value}</div>
            {sub_html}
        </div>""",
        unsafe_allow_html=True,
    )


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
# Header + disclaimer
# ---------------------------------------------------------------------------
st.markdown(
    f"""<div class='main-header'>
        <div class='main-header-icon'>\U0001F9ED</div>
        <div>
            <div class='main-header-title'>Competitor Intelligence Dashboard</div>
            <div class='main-header-sub'>Air Compressor Category &middot; {settings.get('Baseline_Month','May 2026')} vs {settings.get('Report_Month','June 2026')}</div>
        </div>
    </div>""",
    unsafe_allow_html=True,
)
st.markdown(f"<div class='disclaimer-banner'>⚠️ {DISCLAIMER}</div>", unsafe_allow_html=True)

tabs = st.tabs([
    "1. Executive Summary", "2. Competitor Activity", "3. Channel Breakdown",
    "4. Theme & Messaging", "5. Product & Service Signals", "6. Social & Market Conversation",
    "7. PR / News / Events", "8. Hiring & Expansion", "9. Opportunities for Review",
    "10. Raw Data Explorer",
])

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

    # ---- KPI strip ---------------------------------------------------------
    st.markdown("#### \U0001F4CA At a glance")
    r1 = st.columns(4)
    kpi_card(r1[0], "\U0001F3ED", "Competitor activities tracked", exec_metric("Total competitor activities tracked (May+June 2026)"))
    kpi_card(r1[1], "\U0001F3E2", "Competitors tracked", exec_metric("Total competitors tracked"))
    kpi_card(r1[2], "\U0001F4F1", "Social posts (LI+IG+YT)", exec_metric("Total social posts (LinkedIn + Instagram + YouTube)"))
    kpi_card(r1[3], "\U0001F4DD", "Blog posts", exec_metric("Total blog posts"))

    r2 = st.columns(4)
    kpi_card(r2[0], "\U0001F6E0️", "New product/service pages", exec_metric("Total new product/service pages"))
    kpi_card(r2[1], "\U0001F3A4", "Events + webinars", exec_metric("Total events + webinars"))
    kpi_card(r2[2], "\U0001F4F0", "PR releases + news mentions", exec_metric("Total PR releases + news mentions"))
    kpi_card(r2[3], "\U0001F4BC", "LinkedIn jobs / workforce signals", exec_metric("Total LinkedIn jobs / workforce signals"))

    # ---- The Big Picture: 4 synthesis charts -------------------------------
    st.markdown("#### \U0001F50D The big picture")
    cc1, cc2 = st.columns(2)
    with cc1:
        if not overview.empty:
            top_act = overview.sort_values("Total activity", ascending=False)
            fig = px.bar(top_act, x="Total activity", y="Competitor", orientation="h",
                         color="Overall activity level", text="Total activity",
                         color_discrete_map={"High": ATLAS_BLUE, "Medium": ACCENT_AMBER, "Low": NEUTRAL_BAR},
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
    s1, s2, s3 = st.columns(3)
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

    # ---- What to do next: concrete, team-tagged action items ---------------
    st.markdown("#### ✅ What to do next")
    if not opportunities.empty:
        prio_rank = {"High": 0, "Medium": 1, "Low": 2}
        top_opps = opportunities.copy()
        top_opps["_r"] = top_opps["Priority"].map(prio_rank).fillna(3)
        top_opps = top_opps.sort_values("_r").head(5)
        for _, row in top_opps.iterrows():
            teams = [t.strip() for t in str(row["Suggested team to review"]).split(";")]
            tags = "".join(f"<span class='team-tag'>{t}</span>" for t in teams)
            st.markdown(
                f"<div class='next-step'>{tags} {row['Opportunity/Signal']} "
                f"<span style='color:{GRAY};'>&mdash; {row['Why it may matter']}</span></div>",
                unsafe_allow_html=True,
            )
    st.caption("See tab 9 (‘Opportunities for Review’) for the full list with evidence behind each item.")

    st.markdown("---")
    st.caption(DISCLAIMER)

# ===========================================================================
# TAB 2 - COMPETITOR ACTIVITY OVERVIEW
# ===========================================================================
with tabs[1]:
    st.subheader("Competitor Activity Overview")
    st.caption("Activity counts per competitor across all tracked channels, May-June 2026.")
    if not info_if_empty(overview, "competitor activity"):
        view = overview.sort_values("Total activity", ascending=False)

        fig = px.bar(view, x="Competitor", y="Total activity", color="Overall activity level", text="Total activity",
                     color_discrete_map={"High": ATLAS_BLUE, "Medium": ACCENT_AMBER, "Low": NEUTRAL_BAR},
                     title="Total activity by competitor")
        fig.update_traces(textposition="outside", cliponaxis=False, textfont=dict(size=11))
        clean_value_axis(fig, "v")
        st.plotly_chart(style_layout(fig, legend="bottom"), width="stretch")

        leader = view.iloc[0]
        n_high = (view["Overall activity level"] == "High").sum()
        insight(
            f"<b>{leader['Competitor']}</b> is the most active competitor this period with "
            f"{int(leader['Total activity'])} tracked signals. {n_high} of {len(view)} tracked competitors "
            f"are currently in the “High” activity tier, meaning the competitive set overall is fairly vocal right now."
        )
        action(
            "Marketing and social teams may want to keep a closer watch on the highest-activity competitors above, "
            "since they're setting the pace for visibility in this category right now."
        )

        def _level_style(s):
            colors = {"High": ACCENT_RED, "Medium": ACCENT_AMBER, "Low": "#9AA5B1"}
            return [f"background-color:{colors.get(v,'')}; color:white; font-weight:600;" for v in s]

        st.dataframe(
            view.style.apply(_level_style, subset=["Overall activity level"]),
            width="stretch", hide_index=True,
        )

# ===========================================================================
# TAB 3 - CHANNEL ACTIVITY BREAKDOWN
# ===========================================================================
with tabs[2]:
    st.subheader("Channel Activity Breakdown")
    c1, c2 = st.columns(2)
    with c1:
        if not info_if_empty(channel_df, "channel activity"):
            cd_sorted = channel_df.sort_values("Total activity (May+June 2026)", ascending=True)
            cd_sorted = highlight_leader(cd_sorted, "Total activity (May+June 2026)")
            fig = px.bar(cd_sorted, x="Total activity (May+June 2026)", y="Channel", orientation="h",
                         text="Total activity (May+June 2026)", color="_Highlight",
                         color_discrete_map={"Leading": ATLAS_BLUE, "Other": NEUTRAL_BAR}, title="Activity by channel")
            fig.update_yaxes(title="")
            fig.update_traces(textposition="outside", cliponaxis=False, textfont=dict(size=11))
            clean_value_axis(fig, "h")
            st.plotly_chart(style_layout(fig, legend="none"), width="stretch")
    with c2:
        if not info_if_empty(overview, "competitor activity"):
            ov_sorted = overview.sort_values("Total activity", ascending=True)
            ov_sorted = highlight_leader(ov_sorted, "Total activity")
            fig = px.bar(ov_sorted, x="Total activity", y="Competitor", orientation="h", text="Total activity",
                         color="_Highlight", color_discrete_map={"Leading": ATLAS_BLUE_DARK, "Other": NEUTRAL_BAR},
                         title="Activity by competitor")
            fig.update_yaxes(title="")
            fig.update_traces(textposition="outside", cliponaxis=False, textfont=dict(size=11))
            clean_value_axis(fig, "h")
            st.plotly_chart(style_layout(fig, legend="none"), width="stretch")

    if not channel_df.empty:
        cd_sorted = channel_df.sort_values("Total activity (May+June 2026)", ascending=False).reset_index(drop=True)
        total_chan = cd_sorted["Total activity (May+June 2026)"].sum()
        top2 = cd_sorted.head(2)
        insight(
            f"<b>{top2.iloc[0]['Channel']}</b> and <b>{top2.iloc[1]['Channel']}</b> together account for "
            f"{_pct(top2['Total activity (May+June 2026)'].sum(), total_chan)} of all tracked competitor activity. "
            "Most of the competitive noise in this category is happening on social media, not owned channels like blogs or press releases."
        )
        action(
            f"Since {top2.iloc[0]['Channel'].lower()} is where competitors show up most, it's worth prioritizing "
            "monitoring and response cadence there over lower-traffic channels."
        )

    st.markdown("#### Competitor vs. channel (stacked)")
    if not info_if_empty(comp_x_chan, "competitor x channel matrix"):
        chan_cols = [c for c in comp_x_chan.columns if c != "Competitor"]
        long = comp_x_chan.melt(id_vars="Competitor", value_vars=chan_cols, var_name="Channel", value_name="Count")
        long = long[long["Count"] > 0]
        fig = px.bar(long, x="Competitor", y="Count", color="Channel", barmode="stack",
                     color_discrete_sequence=CATEGORICAL_PALETTE, title="Channel mix per competitor")
        st.plotly_chart(style_layout(fig, legend="bottom"), width="stretch")
        with st.expander("View full competitor x channel table"):
            st.dataframe(comp_x_chan, width="stretch", hide_index=True)

    st.markdown("#### May vs. June trend (where data is available)")
    if not info_if_empty(trend_df, "May vs June channel trend"):
        chan_cols = [c for c in trend_df.columns if c != "Month"]
        long = trend_df.melt(id_vars="Month", value_vars=chan_cols, var_name="Channel", value_name="Count")
        long = long[long["Count"] > 0]
        if long.empty:
            st.info("No May 2026 baseline activity was found for these channels - June 2026 figures are shown elsewhere in this dashboard without a comparison point.")
        else:
            fig = px.bar(long, x="Channel", y="Count", color="Month", barmode="group",
                         color_discrete_map={"May 2026": "#9AA5B1", "June 2026": ATLAS_BLUE},
                         title="Channel activity: May 2026 vs June 2026")
            st.plotly_chart(style_layout(fig, legend="bottom"), width="stretch")
        st.caption("May 2026 data is limited to items with a verified publish date; many competitor channels had little "
                    "to no reliably-dated May content in the source data, so May counts may understate true May activity.")

# ===========================================================================
# TAB 4 - THEME & MESSAGING ANALYSIS
# ===========================================================================
with tabs[3]:
    st.subheader("Theme & Messaging Analysis")
    st.caption("Each item is classified into the theme that best matches its content. "
               "\"General brand activity\" covers items that didn't map clearly to a specific theme below.")

    if not info_if_empty(theme_freq, "themes"):
        tf_named = theme_freq[theme_freq["Theme"] != "General brand activity"]
        tf_sorted = highlight_leader(tf_named.sort_values("Frequency", ascending=True), "Frequency")
        fig = px.bar(tf_sorted, x="Frequency", y="Theme", text="Frequency", orientation="h", color="_Highlight",
                     color_discrete_map={"Leading": ATLAS_BLUE, "Other": NEUTRAL_BAR}, title="Top themes by frequency")
        fig.update_traces(textposition="outside", cliponaxis=False, textfont=dict(size=11))
        clean_value_axis(fig, "h")
        st.plotly_chart(style_layout(fig, legend="none"), width="stretch")

        if not tf_named.empty:
            total_named = tf_named["Frequency"].sum()
            top_theme = tf_named.sort_values("Frequency", ascending=False).iloc[0]
            n_themes_multi_brand = None
            if not theme_by_comp.empty and top_theme["Theme"] in theme_by_comp.columns:
                n_themes_multi_brand = (theme_by_comp[top_theme["Theme"]] > 0).sum()
            spread = f" and shows up across {n_themes_multi_brand} of the tracked competitors" if n_themes_multi_brand else ""
            insight(
                f"<b>{top_theme['Theme']}</b> is the single biggest talking point, making up "
                f"{_pct(top_theme['Frequency'], total_named)} of all themed content{spread}. "
                "When a theme is this widespread across the category, it usually signals a shift in what customers are asking about."
            )
            action(
                f"Content, SEO, and product marketing teams may want to check whether our own messaging on "
                f"“{top_theme['Theme']}” is current and visible, since the rest of the category is leaning into it."
            )

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**Themes by competitor**")
        if not info_if_empty(theme_by_comp, "themes by competitor"):
            theme_cols = [c for c in theme_by_comp.columns if c != "Competitor"]
            heat = theme_by_comp.set_index("Competitor")[theme_cols]
            fig = px.imshow(heat, color_continuous_scale=[WHITE, ATLAS_BLUE], aspect="auto",
                             labels=dict(color="Count"), text_auto=True)
            fig.update_layout(margin=dict(l=10, r=10, t=10, b=10))
            st.plotly_chart(style_layout(fig, legend="none"), width="stretch")
    with c2:
        st.markdown("**Themes by channel**")
        if not info_if_empty(theme_by_chan, "themes by channel"):
            theme_cols = [c for c in theme_by_chan.columns if c != "Channel"]
            heat = theme_by_chan.set_index("Channel")[theme_cols]
            fig = px.imshow(heat, color_continuous_scale=[WHITE, ATLAS_BLUE_DARK], aspect="auto",
                             labels=dict(color="Count"), text_auto=True)
            fig.update_layout(margin=dict(l=10, r=10, t=10, b=10))
            st.plotly_chart(style_layout(fig, legend="none"), width="stretch")

    st.markdown("**Theme trend: May vs June 2026**")
    if not info_if_empty(theme_trend, "theme trend"):
        theme_cols = [c for c in theme_trend.columns if c != "Month"]
        long = theme_trend.melt(id_vars="Month", value_vars=theme_cols, var_name="Theme", value_name="Count")
        long = long[long["Count"] > 0]
        fig = px.bar(long, x="Theme", y="Count", color="Month", barmode="group",
                     color_discrete_map={"May 2026": "#9AA5B1", "June 2026": ATLAS_BLUE})
        fig.update_xaxes(tickangle=35)
        st.plotly_chart(style_layout(fig, legend="bottom"), width="stretch")

    with st.expander("View theme-by-competitor and theme-by-channel tables"):
        st.dataframe(theme_by_comp, width="stretch", hide_index=True)
        st.dataframe(theme_by_chan, width="stretch", hide_index=True)

# ===========================================================================
# TAB 5 - PRODUCT & SERVICE SIGNALS
# ===========================================================================
with tabs[4]:
    st.subheader("Product & Service Signals")
    st.caption("New product/service pages, product-related posts, and similar items, with a possible-meaning "
               "read and a suggested team for review.")
    if not info_if_empty(product_signals, "product & service signals"):
        view = product_signals.copy()

        cat_counts = view.groupby("Product/service category").size().reset_index(name="Count")
        cat_counts = highlight_leader(cat_counts.sort_values("Count"), "Count")
        fig = px.bar(cat_counts, x="Count", y="Product/service category", orientation="h",
                     text="Count", color="_Highlight", color_discrete_map={"Leading": ATLAS_BLUE, "Other": NEUTRAL_BAR},
                     title="Signals by product/service category")
        fig.update_yaxes(title="")
        fig.update_traces(textposition="outside", cliponaxis=False, textfont=dict(size=11))
        clean_value_axis(fig, "h")
        st.plotly_chart(style_layout(fig, legend="none"), width="stretch")

        top_cat = cat_counts.sort_values("Count", ascending=False).iloc[0]
        no_dedicated_pages = channel_df.empty or "New product/service pages" not in channel_df["Channel"].values or \
            channel_df.loc[channel_df["Channel"] == "New product/service pages", "Total activity (May+June 2026)"].sum() == 0
        page_note = (
            " Notably, none of this showed up as a dedicated new product/service web page - it's all coming through "
            "social posts and articles instead."
        ) if no_dedicated_pages else ""
        insight(
            f"<b>{top_cat['Product/service category']}</b> is the most common product/service topic competitors are "
            f"posting about, with {int(top_cat['Count'])} signals found.{page_note}"
        )
        action(
            "SEO and product marketing teams may want to check if competitors are under-investing in dedicated product "
            "pages right now — that could be a content gap worth owning, if we move first."
        )

        st.dataframe(
            view[["Competitor", "Product/service category", "Page/post title", "URL", "Date found/published",
                  "Month", "Possible meaning", "Team to review"]],
            width="stretch", hide_index=True,
            column_config={"URL": st.column_config.LinkColumn("URL")},
        )

# ===========================================================================
# TAB 6 - SOCIAL & MARKET CONVERSATION
# ===========================================================================
with tabs[5]:
    st.subheader("Social & Market Conversation")
    st.caption("LinkedIn/social conversation: brand mentions, competitor mentions, \"air compressor\" keyword "
               "mentions, employee posts, and market reaction (comments).")

    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown("**Top conversation themes**")
        if not info_if_empty(conv_top_themes, "conversation themes"):
            ctt = highlight_leader(conv_top_themes.sort_values("Count"), "Count")
            fig = px.bar(ctt, x="Count", y="Theme", orientation="h", text="Count", color="_Highlight",
                         color_discrete_map={"Leading": ATLAS_BLUE, "Other": NEUTRAL_BAR})
            fig.update_yaxes(title="")
            fig.update_traces(textposition="outside", cliponaxis=False, textfont=dict(size=11))
            clean_value_axis(fig, "h")
            st.plotly_chart(style_layout(fig, 360, legend="none"), width="stretch")
    with c2:
        st.markdown("**Sentiment split**")
        if not info_if_empty(conv_sentiment, "sentiment split"):
            fig = px.pie(conv_sentiment, names="Sentiment", values="Count", hole=0.5,
                         color="Sentiment", color_discrete_map=SENTIMENT_COLORS)
            fig.update_traces(textinfo="percent", textfont_size=11)
            st.plotly_chart(style_layout(fig, 360, legend="right"), width="stretch")
    with c3:
        st.markdown("**Most-mentioned brands**")
        if not info_if_empty(brand_mentions, "brand mentions"):
            bm10 = highlight_leader(brand_mentions.sort_values("Mentions").tail(10), "Mentions")
            fig = px.bar(bm10, x="Mentions", y="Brand", orientation="h", text="Mentions", color="_Highlight",
                         color_discrete_map={"Leading": ATLAS_BLUE_DARK, "Other": NEUTRAL_BAR})
            fig.update_yaxes(title="")
            fig.update_traces(textposition="outside", cliponaxis=False, textfont=dict(size=11))
            clean_value_axis(fig, "h")
            st.plotly_chart(style_layout(fig, 360, legend="none"), width="stretch")

    if not conv_sentiment.empty:
        total_sent = conv_sentiment["Count"].sum()
        pos = conv_sentiment.loc[conv_sentiment["Sentiment"] == "Positive", "Count"].sum()
        neg = conv_sentiment.loc[conv_sentiment["Sentiment"] == "Negative", "Count"].sum()
        our_brand_name = str(settings.get("Our_Brand", "Atlas Copco"))
        our_mentions = 0
        if not brand_mentions.empty and our_brand_name in brand_mentions["Brand"].values:
            our_mentions = int(brand_mentions.loc[brand_mentions["Brand"] == our_brand_name, "Mentions"].iloc[0])
        insight(
            f"Conversation in the category is {_pct(pos, total_sent)} positive and {_pct(neg, total_sent)} negative, "
            f"so the overall mood is calm. Our own brand was mentioned {our_mentions} times in this conversation set."
        )
        action(
            "Social and PR teams may want to scan the negative-sentiment posts directly (filterable in tab 10) to "
            "confirm none of them relate to our brand specifically."
        )

    st.markdown("#### Conversation detail")
    if not info_if_empty(social_conv, "social & market conversation"):
        st.dataframe(
            priority_badge_table(social_conv[["Author/company", "Related brand", "Mention type", "Sentiment", "Theme",
                                               "Engagement", "URL", "Short summary", "Possible implication", "Team to review"]]),
            width="stretch", hide_index=True, height=420,
        )

    st.caption("Common pain points / buying themes surface from the \"Top conversation themes\" chart above "
               "(e.g. maintenance/reliability concerns recur most often across brands).")

# ===========================================================================
# TAB 7 - PR, NEWS, EVENTS & WEBINARS
# ===========================================================================
with tabs[6]:
    st.subheader("PR, News, Events & Webinars")
    st.caption("PR releases, news mentions, events, and webinars, grouped together. Items are classified by type: "
               "Brand visibility, Product promotion, Thought leadership, Event participation, Market expansion, "
               "Partnership, or Customer proof/case study.")

    zero_channels = []
    for chan_label in ["Events", "Webinars", "PR releases", "News mentions"]:
        if channel_df.empty or chan_label not in channel_df["Channel"].values or \
           channel_df.loc[channel_df["Channel"] == chan_label, "Total activity (May+June 2026)"].sum() == 0:
            zero_channels.append(chan_label)
    if zero_channels:
        st.warning(
            f"\U0001F50E We found zero confirmed {_join_natural(zero_channels)} from any tracked competitor this period. "
            "That could mean the whole category had a quiet month on these channels, or it could mean our monitoring "
            "isn't catching everything there — either way, it's worth a quick double-check with the PR/Comms team "
            "before assuming nothing happened."
        )

    if not info_if_empty(pr_news, "PR, News, Events & Webinars"):
        view = pr_news.copy()

        fig = px.bar(view.groupby(["Competitor", "Source channel"]).size().reset_index(name="Count"),
                     x="Competitor", y="Count", color="Source channel", barmode="stack",
                     color_discrete_sequence=CATEGORICAL_PALETTE, title="Confirmed items by competitor and channel")
        st.plotly_chart(style_layout(fig, legend="bottom"), width="stretch")

        st.dataframe(
            view[["Competitor", "Source channel", "Type", "Title", "Date", "Month", "URL", "Theme",
                  "Possible meaning", "Team to review"]],
            width="stretch", hide_index=True,
            column_config={"URL": st.column_config.LinkColumn("URL")},
        )

# ===========================================================================
# TAB 8 - HIRING & EXPANSION SIGNALS
# ===========================================================================
with tabs[7]:
    st.subheader("Hiring & Expansion Signals")
    st.caption("LinkedIn jobs / workforce-analytics signals and hiring- or expansion-themed posts. Roles are "
               "grouped into functions such as Sales, Service technician, Engineering, Marketing, Product, "
               "Operations, Customer support, and Regional leadership where the source data indicates a function.")

    if not info_if_empty(hiring, "hiring & expansion signals"):
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("**Signal type mix**")
            fig = px.bar(hiring.groupby(["Competitor", "Signal type"]).size().reset_index(name="Count"),
                         x="Competitor", y="Count", color="Signal type",
                         color_discrete_sequence=[ATLAS_BLUE, ACCENT_AMBER], barmode="stack")
            st.plotly_chart(style_layout(fig, legend="bottom"), width="stretch")
        with c2:
            st.markdown("**Where signals point geographically**")
            loc_counts = hiring["Location (HQ)"].replace("", "Not specified").value_counts().reset_index()
            loc_counts.columns = ["Location", "Count"]
            loc_counts = highlight_leader(loc_counts.sort_values("Count"), "Count")
            fig = px.bar(loc_counts, x="Count", y="Location", orientation="h", text="Count", color="_Highlight",
                         color_discrete_map={"Leading": ATLAS_BLUE_DARK, "Other": NEUTRAL_BAR})
            fig.update_traces(textposition="outside", cliponaxis=False, textfont=dict(size=11))
            clean_value_axis(fig, "h")
            st.plotly_chart(style_layout(fig, legend="none"), width="stretch")

        hc = hiring["Competitor"].value_counts()
        insight(
            f"<b>{hc.index[0]}</b> shows the most hiring & expansion activity, with {int(hc.iloc[0])} of "
            f"{len(hiring)} tracked signals. Rising hiring activity at a competitor can be an early signal of "
            "investment in a specific region, role, or capability."
        )
        action(
            "Talent/HR and regional sales teams may want to take a quick look at which roles and locations are "
            "showing up most below, in case it points to where a competitor is about to ramp up."
        )

        st.dataframe(
            hiring[["Competitor", "Signal type", "Location (HQ)", "Top open-role function (latest snapshot)",
                    "Date", "URL", "Signal detail", "Possible meaning", "Team to review"]],
            width="stretch", hide_index=True,
            column_config={"URL": st.column_config.LinkColumn("URL")},
        )

# ===========================================================================
# TAB 9 - OPPORTUNITIES FOR TEAM REVIEW
# ===========================================================================
with tabs[8]:
    st.subheader("Opportunities for Team Review")
    st.caption("Soft-language signals worth a look. Priority and confidence reflect the strength and reliability "
               "of the underlying evidence, not a recommendation to act.")

    if not info_if_empty(opportunities, "opportunities for review"):
        view = opportunities.copy()

        cc1, cc2 = st.columns(2)
        with cc1:
            prio_counts = view.groupby("Priority").size().reindex(["High", "Medium", "Low"]).dropna().reset_index(name="Count")
            fig = px.bar(prio_counts, x="Priority", y="Count", text="Count", color="Priority", color_discrete_map=PRIORITY_COLORS,
                         title="How many opportunities, by priority?")
            fig.update_traces(textposition="outside", cliponaxis=False, textfont=dict(size=11))
            clean_value_axis(fig, "v")
            st.plotly_chart(style_layout(fig, 320, legend="none"), width="stretch")
        with cc2:
            conf_counts = view.groupby("Confidence").size().reindex(["High", "Medium", "Low"]).dropna().reset_index(name="Count")
            fig2 = px.bar(conf_counts, x="Confidence", y="Count", text="Count", color="Confidence", color_discrete_map=CONFIDENCE_COLORS,
                          title="How confident are we in the evidence?")
            fig2.update_traces(textposition="outside", cliponaxis=False, textfont=dict(size=11))
            clean_value_axis(fig2, "v")
            st.plotly_chart(style_layout(fig2, 320, legend="none"), width="stretch")

        n_high = (view["Priority"] == "High").sum()
        insight(
            f"Of the {len(view)} opportunities surfaced this period, {n_high} are flagged High priority — "
            "meaning the underlying evidence is both strong and clearly tied to a specific competitor action, "
            "not just background noise."
        )

        prio_rank = {"High": 0, "Medium": 1, "Low": 2}
        view["_r"] = view["Priority"].map(prio_rank).fillna(3)
        view = view.sort_values("_r")

        for _, row in view.iterrows():
            p_color = PRIORITY_COLORS.get(row["Priority"], GRAY)
            c_color = CONFIDENCE_COLORS.get(row["Confidence"], GRAY)
            icon = PRIORITY_ICONS.get(row["Priority"], "⚪")
            with st.container(border=True):
                st.markdown(f"**{icon} {row['Opportunity/Signal']}**")
                st.markdown(
                    f"<span style='background:{p_color};color:white;border-radius:10px;padding:2px 10px;font-size:0.78rem;font-weight:600;'>Priority: {row['Priority']}</span> "
                    f"<span style='background:{c_color};color:white;border-radius:10px;padding:2px 10px;font-size:0.78rem;font-weight:600;'>Confidence: {row['Confidence']}</span> "
                    f"<span style='color:{GRAY};font-size:0.82rem;'>Suggested team: {row['Suggested team to review']}</span>",
                    unsafe_allow_html=True,
                )
                st.markdown(f"*Why it may matter:* {row['Why it may matter']}")
                with st.expander("Evidence"):
                    st.write(row["Evidence"])
    st.caption(DISCLAIMER)

# ===========================================================================
# TAB 10 - RAW DATA EXPLORER
# ===========================================================================
with tabs[9]:
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

st.markdown("---")
st.caption(DISCLAIMER)
st.caption("Source: ETL pipeline over LinkedIn, Instagram, YouTube, blog/PR/news scrapes, and LinkedIn workforce "
           "analytics for May-June 2026. See the Settings sheet in the workbook for methodology notes.")
