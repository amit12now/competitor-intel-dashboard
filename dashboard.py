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
ACCENT_RED = "#D8554A"
ACCENT_GREEN = "#3FA66A"

CATEGORICAL_PALETTE = [
    ATLAS_BLUE, ACCENT_AMBER, ACCENT_GREEN, ACCENT_RED, "#7C5CBF",
    "#3D8C97", "#C97FB0", "#5B7FA6", "#A3A86C", "#8D6E63", "#4F6D7A",
]
PRIORITY_COLORS = {"High": ACCENT_RED, "Medium": ACCENT_AMBER, "Low": "#9AA5B1"}
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
.stTabs [data-baseweb="tab-list"] {{ gap: 4px; }}
.stTabs [data-baseweb="tab"] {{
    background-color: {WHITE}; border-radius: 6px 6px 0 0; padding: 8px 14px;
}}
.stTabs [aria-selected="true"] {{ background-color: {ATLAS_BLUE} !important; color: {WHITE} !important; }}
</style>
""", unsafe_allow_html=True)


def style_layout(fig):
    fig.update_layout(
        plot_bgcolor=WHITE, paper_bgcolor=WHITE,
        font=dict(family="Segoe UI, Calibri, sans-serif", color=NAVY, size=12),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
        margin=dict(l=10, r=10, t=40, b=10),
    )
    return fig


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
st.sidebar.markdown("## 🧭 Competitor Intelligence")
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
st.title("Competitor Intelligence Dashboard")
st.caption(f"Air Compressor Category · {settings.get('Baseline_Month','May 2026')} vs {settings.get('Report_Month','June 2026')}")
st.markdown(f"<div class='disclaimer-banner'>⚠️ {DISCLAIMER}</div>", unsafe_allow_html=True)

tabs = st.tabs([
    "1. Executive Summary", "2. Competitor Activity", "3. Channel Breakdown",
    "4. Theme & Messaging", "5. Product & Service Signals", "6. Social & Market Conversation",
    "7. PR / News / Events", "8. Hiring & Expansion", "9. Opportunities for Review",
    "10. Raw Data Explorer",
])

# ===========================================================================
# TAB 1 - EXECUTIVE SUMMARY
# ===========================================================================
with tabs[0]:
    st.subheader("At a glance")
    r1 = st.columns(4)
    r1[0].metric("Competitor activities tracked", exec_metric("Total competitor activities tracked (May+June 2026)"))
    r1[1].metric("Competitors tracked", exec_metric("Total competitors tracked"))
    r1[2].metric("Social posts (LI+IG+YT)", exec_metric("Total social posts (LinkedIn + Instagram + YouTube)"))
    r1[3].metric("Blog posts", exec_metric("Total blog posts"))

    r2 = st.columns(4)
    r2[0].metric("New product/service pages", exec_metric("Total new product/service pages"))
    r2[1].metric("Events + webinars", exec_metric("Total events + webinars"))
    r2[2].metric("PR releases + news mentions", exec_metric("Total PR releases + news mentions"))
    r2[3].metric("LinkedIn jobs / workforce signals", exec_metric("Total LinkedIn jobs / workforce signals"))

    st.markdown("####  ")
    c1, c2 = st.columns([1, 1])
    with c1:
        st.markdown("**Top 3 competitor themes this period**")
        themes_html = "".join(
            f"<span class='theme-pill'>{exec_metric(f'Top theme #{i}')}</span>" for i in (1, 2, 3)
        )
        st.markdown(themes_html, unsafe_allow_html=True)
    with c2:
        st.markdown("**Top 3 possible opportunities for review**")
        for i in (1, 2, 3):
            st.markdown(f"- {exec_metric(f'Possible opportunity #{i}')}")

    st.markdown("#### Summary")
    st.info(exec_metric("Executive summary narrative"))

    st.markdown("#### Where the activity is concentrated")
    cc1, cc2 = st.columns(2)
    with cc1:
        if not overview.empty:
            top_act = overview.sort_values("Total activity", ascending=False)
            fig = px.bar(top_act, x="Total activity", y="Competitor", orientation="h",
                         color="Overall activity level",
                         color_discrete_map={"High": ATLAS_BLUE, "Medium": ACCENT_AMBER, "Low": "#C7CDD4"},
                         title="Total activity by competitor (May+June 2026)")
            fig.update_yaxes(autorange="reversed", title="")
            st.plotly_chart(style_layout(fig), use_container_width=True)
    with cc2:
        if not channel_df.empty:
            fig = px.pie(channel_df, names="Channel", values="Total activity (May+June 2026)", hole=0.45,
                         color_discrete_sequence=CATEGORICAL_PALETTE, title="Activity mix by channel")
            st.plotly_chart(style_layout(fig), use_container_width=True)

    st.caption(DISCLAIMER)

# ===========================================================================
# TAB 2 - COMPETITOR ACTIVITY OVERVIEW
# ===========================================================================
with tabs[1]:
    st.subheader("Competitor Activity Overview")
    st.caption("Activity counts per competitor across all tracked channels, May-June 2026.")
    if not info_if_empty(overview, "competitor activity"):
        level_pick = st.multiselect("Filter by activity level", sorted(overview["Overall activity level"].unique()),
                                     default=sorted(overview["Overall activity level"].unique()), key="act_level")
        view = overview[overview["Overall activity level"].isin(level_pick)] if level_pick else overview

        fig = px.bar(view.sort_values("Total activity", ascending=False),
                     x="Competitor", y="Total activity", color="Overall activity level",
                     color_discrete_map={"High": ATLAS_BLUE, "Medium": ACCENT_AMBER, "Low": "#C7CDD4"},
                     title="Total activity by competitor")
        st.plotly_chart(style_layout(fig), use_container_width=True)

        def _level_style(s):
            colors = {"High": ACCENT_RED, "Medium": ACCENT_AMBER, "Low": "#9AA5B1"}
            return [f"background-color:{colors.get(v,'')}; color:white; font-weight:600;" for v in s]

        st.dataframe(
            view.style.apply(_level_style, subset=["Overall activity level"]),
            use_container_width=True, hide_index=True,
        )

# ===========================================================================
# TAB 3 - CHANNEL ACTIVITY BREAKDOWN
# ===========================================================================
with tabs[2]:
    st.subheader("Channel Activity Breakdown")
    c1, c2 = st.columns(2)
    with c1:
        if not info_if_empty(channel_df, "channel activity"):
            fig = px.bar(channel_df.sort_values("Total activity (May+June 2026)", ascending=True),
                         x="Total activity (May+June 2026)", y="Channel", orientation="h",
                         color_discrete_sequence=[ATLAS_BLUE], title="Activity by channel")
            st.plotly_chart(style_layout(fig), use_container_width=True)
    with c2:
        if not info_if_empty(overview, "competitor activity"):
            fig = px.bar(overview.sort_values("Total activity", ascending=True),
                         x="Total activity", y="Competitor", orientation="h",
                         color_discrete_sequence=[ATLAS_BLUE_DARK], title="Activity by competitor")
            st.plotly_chart(style_layout(fig), use_container_width=True)

    st.markdown("#### Competitor vs. channel (stacked)")
    if not info_if_empty(comp_x_chan, "competitor x channel matrix"):
        chan_cols = [c for c in comp_x_chan.columns if c != "Competitor"]
        long = comp_x_chan.melt(id_vars="Competitor", value_vars=chan_cols, var_name="Channel", value_name="Count")
        long = long[long["Count"] > 0]
        fig = px.bar(long, x="Competitor", y="Count", color="Channel", barmode="stack",
                     color_discrete_sequence=CATEGORICAL_PALETTE, title="Channel mix per competitor")
        st.plotly_chart(style_layout(fig), use_container_width=True)
        with st.expander("View full competitor x channel table"):
            st.dataframe(comp_x_chan, use_container_width=True, hide_index=True)

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
            st.plotly_chart(style_layout(fig), use_container_width=True)
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
        fig = px.bar(theme_freq.sort_values("Frequency", ascending=True), x="Frequency", y="Theme",
                     orientation="h", color_discrete_sequence=[ATLAS_BLUE], title="Top themes by frequency")
        st.plotly_chart(style_layout(fig), use_container_width=True)

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**Themes by competitor**")
        if not info_if_empty(theme_by_comp, "themes by competitor"):
            theme_cols = [c for c in theme_by_comp.columns if c != "Competitor"]
            heat = theme_by_comp.set_index("Competitor")[theme_cols]
            fig = px.imshow(heat, color_continuous_scale=[WHITE, ATLAS_BLUE], aspect="auto",
                             labels=dict(color="Count"))
            fig.update_layout(margin=dict(l=10, r=10, t=10, b=10))
            st.plotly_chart(style_layout(fig), use_container_width=True)
    with c2:
        st.markdown("**Themes by channel**")
        if not info_if_empty(theme_by_chan, "themes by channel"):
            theme_cols = [c for c in theme_by_chan.columns if c != "Channel"]
            heat = theme_by_chan.set_index("Channel")[theme_cols]
            fig = px.imshow(heat, color_continuous_scale=[WHITE, ATLAS_BLUE_DARK], aspect="auto",
                             labels=dict(color="Count"))
            fig.update_layout(margin=dict(l=10, r=10, t=10, b=10))
            st.plotly_chart(style_layout(fig), use_container_width=True)

    st.markdown("**Theme trend: May vs June 2026**")
    if not info_if_empty(theme_trend, "theme trend"):
        theme_cols = [c for c in theme_trend.columns if c != "Month"]
        long = theme_trend.melt(id_vars="Month", value_vars=theme_cols, var_name="Theme", value_name="Count")
        long = long[long["Count"] > 0]
        fig = px.bar(long, x="Theme", y="Count", color="Month", barmode="group",
                     color_discrete_map={"May 2026": "#9AA5B1", "June 2026": ATLAS_BLUE})
        fig.update_xaxes(tickangle=35)
        st.plotly_chart(style_layout(fig), use_container_width=True)

    with st.expander("View theme-by-competitor and theme-by-channel tables"):
        st.dataframe(theme_by_comp, use_container_width=True, hide_index=True)
        st.dataframe(theme_by_chan, use_container_width=True, hide_index=True)

# ===========================================================================
# TAB 5 - PRODUCT & SERVICE SIGNALS
# ===========================================================================
with tabs[4]:
    st.subheader("Product & Service Signals")
    st.caption("New product/service pages, product-related posts, and similar items, with a possible-meaning "
               "read and a suggested team for review.")
    if not info_if_empty(product_signals, "product & service signals"):
        cats = sorted(product_signals["Product/service category"].dropna().unique())
        comps = sorted(product_signals["Competitor"].dropna().unique())
        f1, f2 = st.columns(2)
        cat_pick = f1.multiselect("Filter by category", cats, default=[], key="prod_cat")
        comp_pick = f2.multiselect("Filter by competitor", comps, default=[], key="prod_comp")
        view = product_signals.copy()
        if cat_pick:
            view = view[view["Product/service category"].isin(cat_pick)]
        if comp_pick:
            view = view[view["Competitor"].isin(comp_pick)]

        cat_counts = view.groupby("Product/service category").size().reset_index(name="Count")
        fig = px.bar(cat_counts.sort_values("Count"), x="Count", y="Product/service category", orientation="h",
                     color_discrete_sequence=[ATLAS_BLUE], title="Signals by product/service category")
        fig.update_yaxes(title="")
        st.plotly_chart(style_layout(fig), use_container_width=True)

        st.dataframe(
            view[["Competitor", "Product/service category", "Page/post title", "URL", "Date found/published",
                  "Month", "Possible meaning", "Team to review"]],
            use_container_width=True, hide_index=True,
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
            fig = px.bar(conv_top_themes.sort_values("Count"), x="Count", y="Theme", orientation="h",
                         color_discrete_sequence=[ATLAS_BLUE])
            fig.update_yaxes(title="")
            st.plotly_chart(style_layout(fig), use_container_width=True)
    with c2:
        st.markdown("**Sentiment split**")
        if not info_if_empty(conv_sentiment, "sentiment split"):
            fig = px.pie(conv_sentiment, names="Sentiment", values="Count", hole=0.45,
                         color="Sentiment", color_discrete_map=SENTIMENT_COLORS)
            st.plotly_chart(style_layout(fig), use_container_width=True)
    with c3:
        st.markdown("**Most-mentioned brands**")
        if not info_if_empty(brand_mentions, "brand mentions"):
            fig = px.bar(brand_mentions.sort_values("Mentions").tail(10), x="Mentions", y="Brand", orientation="h",
                         color_discrete_sequence=[ATLAS_BLUE_DARK])
            fig.update_yaxes(title="")
            st.plotly_chart(style_layout(fig), use_container_width=True)

    st.markdown("#### Conversation detail")
    if not info_if_empty(social_conv, "social & market conversation"):
        f1, f2, f3 = st.columns(3)
        mt_pick = f1.multiselect("Mention type", sorted(social_conv["Mention type"].dropna().unique()), key="conv_mt")
        sent_pick = f2.multiselect("Sentiment", sorted(social_conv["Sentiment"].dropna().unique()), key="conv_sent")
        brand_pick = f3.multiselect("Related brand", sorted(social_conv["Related brand"].dropna().unique()), key="conv_brand")
        view = social_conv.copy()
        if mt_pick:
            view = view[view["Mention type"].isin(mt_pick)]
        if sent_pick:
            view = view[view["Sentiment"].isin(sent_pick)]
        if brand_pick:
            view = view[view["Related brand"].isin(brand_pick)]
        st.dataframe(
            priority_badge_table(view[["Author/company", "Related brand", "Mention type", "Sentiment", "Theme",
                                        "Engagement", "URL", "Short summary", "Possible implication", "Team to review"]]),
            use_container_width=True, hide_index=True,
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
            "No signals found this period for: " + ", ".join(zero_channels) + ". "
            "This may reflect limited competitor activity in these channels in May-June 2026, or a gap in "
            "monitoring coverage for these specific channels - it's worth validating directly with the PR/Comms "
            "team before concluding no activity occurred. (See Opportunities for Review and Settings/methodology "
            "notes for more detail.)"
        )

    if not info_if_empty(pr_news, "PR, News, Events & Webinars"):
        f1, f2 = st.columns(2)
        type_pick = f1.multiselect("Filter by type", sorted(pr_news["Type"].dropna().unique()), key="pr_type")
        comp_pick = f2.multiselect("Filter by competitor", sorted(pr_news["Competitor"].dropna().unique()), key="pr_comp")
        view = pr_news.copy()
        if type_pick:
            view = view[view["Type"].isin(type_pick)]
        if comp_pick:
            view = view[view["Competitor"].isin(comp_pick)]

        fig = px.bar(view.groupby(["Competitor", "Source channel"]).size().reset_index(name="Count"),
                     x="Competitor", y="Count", color="Source channel", barmode="stack",
                     color_discrete_sequence=CATEGORICAL_PALETTE, title="Confirmed items by competitor and channel")
        st.plotly_chart(style_layout(fig), use_container_width=True)

        st.dataframe(
            view[["Competitor", "Source channel", "Type", "Title", "Date", "Month", "URL", "Theme",
                  "Possible meaning", "Team to review"]],
            use_container_width=True, hide_index=True,
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
            st.plotly_chart(style_layout(fig), use_container_width=True)
        with c2:
            st.markdown("**Where signals point geographically**")
            loc_counts = hiring["Location (HQ)"].replace("", "Not specified").value_counts().reset_index()
            loc_counts.columns = ["Location", "Count"]
            fig = px.bar(loc_counts.sort_values("Count"), x="Count", y="Location", orientation="h",
                         color_discrete_sequence=[ATLAS_BLUE_DARK])
            st.plotly_chart(style_layout(fig), use_container_width=True)

        comp_pick = st.multiselect("Filter by competitor", sorted(hiring["Competitor"].dropna().unique()), key="hiring_comp")
        view = hiring[hiring["Competitor"].isin(comp_pick)] if comp_pick else hiring
        st.dataframe(
            view[["Competitor", "Signal type", "Location (HQ)", "Top open-role function (latest snapshot)",
                  "Date", "URL", "Signal detail", "Possible meaning", "Team to review"]],
            use_container_width=True, hide_index=True,
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
        f1, f2 = st.columns(2)
        prio_pick = f1.multiselect("Filter by priority", ["High", "Medium", "Low"], default=[], key="opp_prio")
        team_pick = f2.multiselect("Filter by suggested team", sorted(opportunities["Suggested team to review"].dropna().unique()), key="opp_team")
        view = opportunities.copy()
        if prio_pick:
            view = view[view["Priority"].isin(prio_pick)]
        if team_pick:
            view = view[view["Suggested team to review"].isin(team_pick)]

        cc1, cc2 = st.columns(2)
        with cc1:
            prio_counts = view.groupby("Priority").size().reindex(["High", "Medium", "Low"]).dropna().reset_index(name="Count")
            fig = px.bar(prio_counts, x="Priority", y="Count", color="Priority", color_discrete_map=PRIORITY_COLORS)
            st.plotly_chart(style_layout(fig), use_container_width=True)
        with cc2:
            fig2 = px.bar(view.groupby("Confidence").size().reindex(["High", "Medium", "Low"]).dropna().reset_index(name="Count"),
                          x="Confidence", y="Count", color="Confidence", color_discrete_map=CONFIDENCE_COLORS)
            st.plotly_chart(style_layout(fig2), use_container_width=True)

        for _, row in view.iterrows():
            p_color = PRIORITY_COLORS.get(row["Priority"], GRAY)
            c_color = CONFIDENCE_COLORS.get(row["Confidence"], GRAY)
            with st.container(border=True):
                st.markdown(f"**{row['Opportunity/Signal']}**")
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
            use_container_width=True, hide_index=True, height=520,
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
