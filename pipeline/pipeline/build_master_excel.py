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
for metric_col in ["likes", "comments", "shares", "reactions", "views"]:
    if metric_col not in df.columns:
        df[metric_col] = 0
    df[metric_col] = pd.to_numeric(df[metric_col], errors="coerce").fillna(0).astype(int)
if "content_format" not in df.columns:
    df["content_format"] = ""

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
# June dashboard-ready summary sheets
# ---------------------------------------------------------------------------
june = df[df["month"] == "June 2026"].copy()
social_post_mask = june["channel"].isin(["LinkedIn", "Instagram", "YouTube", "Employee posts"]) & (june["mention_type"] != "Comment")

def _top_theme(sub):
    themes = sub.loc[sub["theme_primary"] != "General brand activity", "theme_primary"]
    if themes.empty:
        return "General brand activity" if not sub.empty else "No signals found this period"
    return themes.value_counts().index[0]

def _latest_signal(sub):
    dated = sub.sort_values("date", ascending=False)
    return dated["date"].iloc[0] if not dated.empty else ""

brief_rows = []
for comp in TRACKED_COMPETITORS:
    sub = june[june["competitor"] == comp]
    social = sub[sub["channel"].isin(["LinkedIn", "Instagram", "YouTube", "Employee posts"]) & (sub["mention_type"] != "Comment")]
    training = sub[sub["theme_primary"] == "Training / webinar"]
    brief_rows.append({
        "Competitor": comp,
        "Total June signals": int(len(sub)),
        "Social posts": int(len(social)),
        "Social engagement": int(social["engagement"].sum()),
        "Avg social engagement/post": round(float(social["engagement"].mean()), 1) if len(social) else 0,
        "LinkedIn posts": int(((sub["channel"] == "LinkedIn") & (sub["mention_type"] != "Comment")).sum()),
        "Instagram posts": int((sub["channel"] == "Instagram").sum()),
        "YouTube videos": int((sub["channel"] == "YouTube").sum()),
        "Blog posts": int((sub["channel"] == "Blog").sum()),
        "Confirmed webinar listings": int((sub["channel"] == "Webinars").sum()),
        "Training/webinar-themed signals": int(len(training)),
        "Confirmed PR releases": int((sub["channel"] == "PR").sum()),
        "News items": int((sub["channel"] == "News").sum()),
        "Hiring/workforce signals": int((sub["channel"] == "Jobs").sum()),
        "High-priority signals": int((sub["priority"] == "High").sum()),
        "Top non-general theme": _top_theme(sub),
        "Latest signal date": _latest_signal(sub),
    })
brief = pd.DataFrame(brief_rows).sort_values("Total June signals", ascending=False).reset_index(drop=True)
SHEETS["Clean_June_Competitor_Brief"] = brief

content_mix_rows = []
for comp in TRACKED_COMPETITORS:
    sub = june[june["competitor"] == comp]
    content_mix_rows.append({
        "Competitor": comp,
        "Social media posts": int((sub["channel"].isin(["LinkedIn", "Instagram", "YouTube", "Employee posts"]) & (sub["mention_type"] != "Comment")).sum()),
        "Blog posts": int((sub["channel"] == "Blog").sum()),
        "Confirmed webinar listings": int((sub["channel"] == "Webinars").sum()),
        "Training/webinar-themed signals": int((sub["theme_primary"] == "Training / webinar").sum()),
        "Confirmed PR releases": int((sub["channel"] == "PR").sum()),
        "Google News / news items": int((sub["channel"] == "News").sum()),
        "Product/service signals": int((sub["product_category"] != "General / Brand").sum()),
        "Hiring/workforce signals": int((sub["channel"] == "Jobs").sum()),
    })
content_mix = pd.DataFrame(content_mix_rows).sort_values("Social media posts", ascending=False).reset_index(drop=True)
SHEETS["Clean_June_Content_Mix"] = content_mix

social_rows = []
social_df = june.loc[social_post_mask].copy()
for (comp, channel), sub in social_df.groupby(["competitor", "channel"]):
    top = sub.sort_values("engagement", ascending=False).head(1)
    top_title = top["title"].iloc[0] if len(top) else ""
    top_url = top["url"].iloc[0] if len(top) else ""
    social_rows.append({
        "Competitor": comp,
        "Platform": channel,
        "Posts/videos": int(len(sub)),
        "Likes": int(sub["likes"].sum()),
        "Comments": int(sub["comments"].sum()),
        "Shares": int(sub["shares"].sum()),
        "Reactions": int(sub["reactions"].sum()),
        "Views": int(sub["views"].sum()),
        "Engagement": int(sub["engagement"].sum()),
        "Avg engagement/post": round(float(sub["engagement"].mean()), 1) if len(sub) else 0,
        "Top post/video": top_title,
        "Top post/video URL": top_url,
    })
social_metrics = pd.DataFrame(social_rows)
if not social_metrics.empty:
    social_metrics = social_metrics.sort_values(["Engagement", "Posts/videos"], ascending=False).reset_index(drop=True)
SHEETS["Clean_June_Social_Metrics"] = social_metrics

def _content_type(row):
    if row["channel"] == "Blog":
        return "Blog post"
    if row["channel"] == "Webinars":
        return "Confirmed webinar listing"
    if row["theme_primary"] == "Training / webinar":
        return "Training/webinar-themed signal"
    if row["channel"] == "PR":
        return "Confirmed PR release"
    if row["channel"] == "News":
        return "News item"
    if row["channel"] in ["LinkedIn", "Instagram", "YouTube", "Employee posts"] and row["mention_type"] != "Comment":
        return "Social media post"
    if row["channel"] == "Jobs":
        return "Hiring/workforce signal"
    return row["channel"]

content_items = june.copy()
content_items["Dashboard content type"] = content_items.apply(_content_type, axis=1)
content_items = content_items[[
    "date", "competitor", "Dashboard content type", "channel", "source_label",
    "title", "url", "theme_primary", "sentiment", "engagement", "likes",
    "comments", "shares", "views", "priority", "confidence", "possible_meaning",
    "team_to_review",
]].rename(columns={
    "date": "Date", "competitor": "Competitor", "channel": "Channel",
    "source_label": "Source type", "title": "Title", "url": "URL",
    "theme_primary": "Theme", "sentiment": "Sentiment", "engagement": "Engagement",
    "likes": "Likes", "comments": "Comments", "shares": "Shares", "views": "Views",
    "priority": "Priority", "confidence": "Confidence",
    "possible_meaning": "Possible meaning", "team_to_review": "Team to review",
})
content_items = content_items.sort_values(["Date", "Engagement"], ascending=[False, False]).reset_index(drop=True)
SHEETS["Clean_June_Content_Items"] = content_items

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
