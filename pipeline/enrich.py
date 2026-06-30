# -*- coding: utf-8 -*-
"""
Stage 4: enrich the scoped raw records with theme, sentiment, geography,
product/service category, possible meaning, team-to-review, priority and
confidence -- all rule-based / template-driven from the actual record text
and metadata (no fabricated facts).
"""
import pandas as pd
from taxonomy import classify_themes, classify_sentiment, find_geography, THEME_INSIGHT

df = pd.read_csv("stage3_scoped.csv", keep_default_na=False)
df["content"] = df["content"].fillna("")
df["title"] = df["title"].fillna("")

text_for_theme = (df["title"] + " " + df["content"]).str.slice(0, 2000)

themes = text_for_theme.apply(classify_themes)
df["theme_primary"] = themes.apply(lambda lst: lst[0] if lst else "General brand activity")
df["theme_secondary"] = themes.apply(lambda lst: lst[1] if len(lst) > 1 else "")
df["themes_all"] = themes.apply(lambda lst: "; ".join(lst) if lst else "General brand activity")

df["sentiment"] = text_for_theme.apply(classify_sentiment)
df["geography"] = text_for_theme.apply(find_geography)

PRODUCT_CATEGORY_MAP = {
    "Oil-free compressors": "Oil-free compressor",
    "Rotary screw compressors": "Rotary screw compressor",
    "Service / aftermarket": "Aftermarket / Service",
    "Maintenance / repair": "Maintenance / Repair",
    "Rental": "Rental / Fleet",
    "New product launch": "New product / service",
    "Dealer / distributor": "Dealer / Distributor network",
    "Training / webinar": "Training / Education",
    "Sustainability": "Sustainability / ESG",
    "Energy efficiency": "Energy-efficient systems",
    "Customer success / case study": "Customer success story",
    "Hiring / expansion": "Hiring / Facility expansion",
    "Local market activity": "Events / Market activity",
    "Industrial productivity": "Industrial productivity",
}
df["product_category"] = df["theme_primary"].map(PRODUCT_CATEGORY_MAP).fillna("General / Brand")

df["possible_meaning"] = df["theme_primary"].map(lambda t: THEME_INSIGHT.get(t, THEME_INSIGHT["General brand activity"])[0])
df["team_to_review"] = df["theme_primary"].map(lambda t: THEME_INSIGHT.get(t, THEME_INSIGHT["General brand activity"])[1])

# Channel-specific overrides for team_to_review / possible_meaning where the
# channel itself is the strongest signal (jobs, employee posts, brand mentions)
def refine_row(row):
    meaning, team = row["possible_meaning"], row["team_to_review"]
    if row["channel"] == "Jobs":
        meaning = "Hiring/workforce trend signal -- " + meaning
        team = "Leadership / Sales / Regional Marketing"
    elif row["mention_type"] == "Employee Post":
        meaning = "Employee/associated-person commentary; may reflect informal brand sentiment or internal culture signals worth monitoring."
        team = "Marketing / PR"
    elif row["mention_type"] == "Our Brand Mention":
        meaning = "Third-party mention of our brand; relevant for brand sentiment and share-of-voice monitoring."
        team = "PR / Marketing"
    elif row["mention_type"] == "Comment":
        meaning = "Market reaction/comment on a competitor post; may offer a read on audience sentiment or pain points."
        team = "Social / PR"
    elif row["mention_type"] == "Air Compressor Keyword" and row["competitor"] == "Unspecified / General":
        meaning = "General industry conversation not tied to a tracked competitor; useful as market-context signal only."
        team = "Marketing (context only)"
    return pd.Series([meaning, team])

df[["possible_meaning", "team_to_review"]] = df.apply(refine_row, axis=1)

# Confidence: how much we trust this record's attribution / dating
CONF_BY_MENTION = {
    "Company Post": "High", "Hiring Signal": "High",
    "Employee Post": "Medium", "Competitor Mention": "Medium",
    "Our Brand Mention": "Medium", "Comment": "Low", "Air Compressor Keyword": "Low",
}
df["confidence"] = df["mention_type"].map(CONF_BY_MENTION).fillna("Medium")

# Priority: combine theme weight + engagement (relative to channel median) + mention type
HIGH_WEIGHT_THEMES = {"New product launch", "Hiring / expansion", "Sustainability", "Service / aftermarket"}
chan_median_eng = df.groupby("channel")["engagement"].median().to_dict()

def priority_row(row):
    score = 0
    if row["theme_primary"] in HIGH_WEIGHT_THEMES:
        score += 2
    if row["mention_type"] in ("Company Post", "Hiring Signal"):
        score += 1
    med = chan_median_eng.get(row["channel"], 0)
    if med and row["engagement"] > med:
        score += 1
    if row["mention_type"] in ("Comment", "Air Compressor Keyword") and row["competitor"] == "Unspecified / General":
        score -= 2
    if score >= 3:
        return "High"
    if score >= 1:
        return "Medium"
    return "Low"

df["priority"] = df.apply(priority_row, axis=1)

df.to_csv("stage4_enriched.csv", index=False)
print("Enriched rows:", len(df))
print()
print("Theme distribution:")
print(df["theme_primary"].value_counts())
print()
print("Sentiment distribution:")
print(df["sentiment"].value_counts())
print()
print("Priority distribution:")
print(df["priority"].value_counts())
print()
print("Confidence distribution:")
print(df["confidence"].value_counts())
print()
print("Geography non-empty count:", (df["geography"] != "").sum())
