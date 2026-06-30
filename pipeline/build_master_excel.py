# -*- coding: utf-8 -*-
"""
Stage 5: build the master Excel workbook (all dashboard-ready "Clean_*" sheets,
plus Settings) from stage4_enriched.csv. dashboard.py reads ONLY from this file.
"""
import pandas as pd
import numpy as np
from datetime import datetime

df = pd.read_csv("stage4_enriched.csv", keep_default_na=False)
df["engagement"] = pd.to_numeric(df["engagement"], errors="coerce").fillna(0).astype(int)

TRACKED_COMPETITORS = [
    "Ingersoll Rand", "Sullair", "ELGi", "FS-Curtis", "Quincy Compressor",
    "Kaeser Compressors", "BOGE", "Hitachi Global Air Power", "Gardner Denver",
]
OUR_BRAND = "Atlas Copco"

CHANNEL_TO_COLUMN = {
    "LinkedIn": "LinkedIn posts", "Instagram": "Instagram posts", "YouTube": "YouTube videos",
    "Blog": "Blogs", "Product pages": "New product/service pages", "Events": "Events",
    "Webinars": "Webinars", "PR": "PR releases", "News": "News mentions",
    "Jobs": "LinkedIn jobs", "Employee posts": "Employee/person posts",
}
ALL_CHANNELS = list(CHANNEL_TO_COLUMN.keys())

SHEETS = {}

# ---------------------------------------------------------------------------
# Clean_Competitor_Activity
# ---------------------------------------------------------------------------
rows = []
for comp in TRACKED_COMPETITORS:
    sub = df[df["competitor"] == comp]
    counts = {}
    counts["LinkedIn posts"] = int(((sub["channel"] == "LinkedIn") & (sub["mention_type"] != "Comment")).sum())
    counts["Instagram posts"] = int((sub["channel"] == "Instagram").sum())
    counts["YouTube videos"] = int((sub["channel"] == "YouTube").sum())
    counts["Blogs"] = int((sub["channel"] == "Blog").sum())
    counts["New product/service pages"] = int((sub["channel"] == "Product pages").sum())
    counts["Events"] = int((sub["channel"] == "Events").sum())
    counts["Webinars"] = int((sub["channel"] == "Webinars").sum())
    counts["PR releases"] = int((sub["channel"] == "PR").sum())
    counts["News mentions"] = int((sub["channel"] == "News").sum())
    counts["LinkedIn jobs"] = int((sub["channel"] == "Jobs").sum())
    counts["Employee/person posts"] = int((sub["channel"] == "Employee posts").sum())
    total = sum(counts.values())
    rows.append({"Competitor": comp, **counts, "Total activity": total})

overview = pd.DataFrame(rows)
hi_cut = overview["Total activity"].quantile(0.66)
lo_cut = overview["Total activity"].quantile(0.33)
def activity_level(t):
    if t >= hi_cut and t >= 40:
        return "High"
    if t <= lo_cut or t < 20:
        return "Low"
    return "Medium"
overview["Overall activity level"] = overview["Total activity"].apply(activity_level)
overview = overview.sort_values("Total activity", ascending=False).reset_index(drop=True)
SHEETS["Clean_Competitor_Activity"] = overview

# ---------------------------------------------------------------------------
# Clean_Channel_Activity / Clean_Competitor_x_Channel / Clean_Trend_May_June
# ---------------------------------------------------------------------------
chan_counts = df["channel"].value_counts().reindex(ALL_CHANNELS, fill_value=0)
channel_breakdown = pd.DataFrame({
    "Channel": [CHANNEL_TO_COLUMN[c] for c in chan_counts.index],
    "Total activity (May+June 2026)": chan_counts.values,
}).sort_values("Total activity (May+June 2026)", ascending=False).reset_index(drop=True)
SHEETS["Clean_Channel_Activity"] = channel_breakdown

pivot = pd.crosstab(df["competitor"], df["channel"])
idx_order = TRACKED_COMPETITORS + [OUR_BRAND, "Unspecified / General"]
pivot = pivot.reindex(index=[i for i in idx_order if i in pivot.index])
pivot = pivot[[c for c in ALL_CHANNELS if c in pivot.columns]]
pivot = pivot.rename(columns=CHANNEL_TO_COLUMN)
pivot.index.name = "Competitor"
SHEETS["Clean_Competitor_x_Channel"] = pivot.reset_index()

trend = pd.crosstab(df["month"], df["channel"]).reindex(index=["May 2026", "June 2026"], fill_value=0)
trend = trend.rename(columns=CHANNEL_TO_COLUMN)
SHEETS["Clean_Trend_May_vs_June"] = trend.reset_index().rename(columns={"month": "Month"})

# ---------------------------------------------------------------------------
# Clean_Theme_* sheets
# ---------------------------------------------------------------------------
theme_freq = df["theme_primary"].value_counts().reset_index()
theme_freq.columns = ["Theme", "Frequency"]
SHEETS["Clean_Theme_Frequency"] = theme_freq

theme_by_competitor = pd.crosstab(df["competitor"], df["theme_primary"])
SHEETS["Clean_Theme_by_Competitor"] = theme_by_competitor.reset_index().rename(columns={"competitor": "Competitor"})

theme_by_channel = pd.crosstab(df["channel"], df["theme_primary"])
SHEETS["Clean_Theme_by_Channel"] = theme_by_channel.reset_index().rename(columns={"channel": "Channel"})

theme_trend = pd.crosstab(df["month"], df["theme_primary"]).reindex(index=["May 2026", "June 2026"], fill_value=0)
SHEETS["Clean_Theme_Trend_May_vs_June"] = theme_trend.reset_index().rename(columns={"month": "Month"})

# ---------------------------------------------------------------------------
# Clean_Product_Service_Signals
# ---------------------------------------------------------------------------
prod_mask = df["product_category"] != "General / Brand"
prod = df.loc[prod_mask, [
    "competitor", "product_category", "title", "url", "date", "month", "possible_meaning", "team_to_review",
]].rename(columns={
    "competitor": "Competitor", "product_category": "Product/service category",
    "title": "Page/post title", "url": "URL", "date": "Date found/published", "month": "Month",
    "possible_meaning": "Possible meaning", "team_to_review": "Team to review",
})
SHEETS["Clean_Product_Service_Signals"] = prod.reset_index(drop=True)

# ---------------------------------------------------------------------------
# Clean_Social_Conversation + summaries
# ---------------------------------------------------------------------------
conv_mask = df["mention_type"].isin(["Competitor Mention", "Our Brand Mention", "Air Compressor Keyword", "Employee Post", "Comment"]) & (df["channel"] != "News")
conv = df.loc[conv_mask, [
    "author_name", "competitor", "mention_type", "sentiment", "theme_primary", "engagement",
    "url", "content", "possible_meaning", "team_to_review",
]].copy()
conv["content"] = conv["content"].str.slice(0, 220)
conv = conv.rename(columns={
    "author_name": "Author/company", "competitor": "Related brand", "mention_type": "Mention type",
    "sentiment": "Sentiment", "theme_primary": "Theme", "engagement": "Engagement", "url": "URL",
    "content": "Short summary", "possible_meaning": "Possible implication", "team_to_review": "Team to review",
})
conv = conv.sort_values("Engagement", ascending=False).reset_index(drop=True)
SHEETS["Clean_Social_Conversation"] = conv

conv_theme_top = conv["Theme"].value_counts().head(10).reset_index()
conv_theme_top.columns = ["Theme", "Count"]
SHEETS["Clean_Conversation_Top_Themes"] = conv_theme_top

sentiment_split = conv["Sentiment"].value_counts().reset_index()
sentiment_split.columns = ["Sentiment", "Count"]
SHEETS["Clean_Conversation_Sentiment"] = sentiment_split

brand_mentions = conv["Related brand"].value_counts().reset_index()
brand_mentions.columns = ["Brand", "Mentions"]
SHEETS["Clean_Most_Mentioned_Brands"] = brand_mentions

# ---------------------------------------------------------------------------
# Clean_PR_News_Events_Webinars  (honest: will be empty / near-empty)
# ---------------------------------------------------------------------------
PR_TYPE_MAP = {
    "Local market activity": "Event participation", "Training / webinar": "Thought leadership",
    "New product launch": "Product promotion", "Hiring / expansion": "Market expansion",
    "Dealer / distributor": "Partnership", "Customer success / case study": "Customer proof / case study",
}
pr_mask = df["channel"].isin(["PR", "News", "Events", "Webinars", "Blog"])
pr = df.loc[pr_mask].copy()
pr["Type"] = pr["theme_primary"].map(PR_TYPE_MAP).fillna("Brand visibility")
pr_out = pr[["competitor", "channel", "Type", "title", "date", "month", "url", "theme_primary", "possible_meaning", "team_to_review"]].rename(columns={
    "competitor": "Competitor", "channel": "Source channel", "title": "Title", "date": "Date", "month": "Month",
    "url": "URL", "theme_primary": "Theme", "possible_meaning": "Possible meaning", "team_to_review": "Team to review",
})
SHEETS["Clean_PR_News_Events"] = pr_out.reset_index(drop=True)

# ---------------------------------------------------------------------------
# Clean_Hiring_Expansion_Signals
# ---------------------------------------------------------------------------
jobs_mask = df["channel"] == "Jobs"
jobs = df.loc[jobs_mask].copy()
hiring_qual_mask = (df["theme_primary"] == "Hiring / expansion") & (~jobs_mask)
hiring_qual = df.loc[hiring_qual_mask].copy()

jobs_out = jobs[["competitor", "job_location", "job_top_function", "date", "url", "content", "possible_meaning", "team_to_review"]].rename(columns={
    "competitor": "Competitor", "job_location": "Location (HQ)", "job_top_function": "Top open-role function (latest snapshot)",
    "date": "Date", "url": "URL", "content": "Signal detail", "possible_meaning": "Possible meaning", "team_to_review": "Team to review",
})
jobs_out.insert(1, "Signal type", "Workforce / hiring trend (aggregate analytics)")

hiring_qual_out = hiring_qual[["competitor", "geography", "date", "url", "title", "possible_meaning", "team_to_review"]].rename(columns={
    "competitor": "Competitor", "geography": "Location (HQ)", "date": "Date", "url": "URL",
    "title": "Signal detail", "possible_meaning": "Possible meaning", "team_to_review": "Team to review",
})
hiring_qual_out.insert(1, "Signal type", "Hiring / expansion post (social)")
hiring_qual_out["Top open-role function (latest snapshot)"] = ""

hiring_combined = pd.concat([jobs_out, hiring_qual_out], ignore_index=True, sort=False)
col_order = ["Competitor", "Signal type", "Location (HQ)", "Top open-role function (latest snapshot)", "Date", "URL", "Signal detail", "Possible meaning", "Team to review"]
hiring_combined = hiring_combined[col_order]
SHEETS["Clean_Hiring_Expansion"] = hiring_combined

print("Stage-1 sheets built OK:", len(SHEETS))
import pickle
with open("sheets_stage1.pkl", "wb") as f:
    pickle.dump((df, SHEETS), f)
