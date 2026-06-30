# -*- coding: utf-8 -*-
import json
from datetime import datetime
import pandas as pd
from taxonomy import classify_themes, classify_sentiment, THEME_INSIGHT, ALIAS_TO_CANON, OUR_BRAND_ALIASES

SRC = "dataset_google-news-scraper-fast_2026-06-30_04-31-13-598.json"

with open(SRC) as f:
    data = json.load(f)

KEEP_KW = {
    "Industrial Air compressor", "Air compressor", "Atlas Copco", "Ingersoll Rand",
    "Elgi compressor", "sullair", "Quincy Compressor", "Garden Denver",
    "Rotary screw air compressor", "Oil free air compressor",
    "Industrial Air dryer", "Industrial Air Chiller",
}

def in_window(dt_str):
    try:
        dt = datetime.fromisoformat(dt_str.replace("Z", "+00:00"))
    except Exception:
        return False
    return dt.year == 2026 and dt.month in (5, 6)

EXCLUDE_TITLES = {
    "The sun is back, and so are two of Denver’s biggest spring plant sales",
    "Unvaccinated Denver singles looking for new bar after Recess event canceled",
    "5 Tips for Shopping the Denver Botanic Gardens Plant Sale Like a Pro",
    "Creating fairy gardens spark imagination for kids and adults alike",
    "The 10 Denver beer gardens we’re loving now",
    "Denver high school for pregnant, parenting teens cuts ribbon on new community garden",
    "Proctor's Garden Tour: Explore an expert garden in Denver",
    "Denver Botanic Gardens’ summer exhibit offers a chance to introduce kids to fine art",
    "A 12-foot tall plant at Denver Botanic Gardens is in its 'death bloom'",
    "Agave plant enters dramatic ‘death bloom’ at Denver Botanic Gardens",
    "Rick Rodgers Obituary (2026) - Quincy, IL - Hansen-Spear Funeral Home - Quincy",
    "Steve Dietrich — Lummis-Hamilton Funeral Home",
    "Spotlight on Baldwin County",
    "Spotlight on Baldwin County: Economic Engines",
    "Ronald Holtshouser Obituary (1951 - 2026) - Quincy, IL - WGEM",
    "Ronald Holtshouser Obituary (2026) - Quincy, IL - Duker & Haugh Funeral Home",
    "The 4 Best Hair Dryers of 2026 | Reviews by Wirecutter",
    "Somerset Fire Crews extinguish industrial building fire",
    "Garage Like An Oven? Keep The Cool Air Flowing With These Fans",
    "The 7 Best Professional Hair Dryers That Help Me Achieve a Salon-Quality Blowout at Home",
    "The Autopilot for Industry: Autonomous Process Control and the Future of Operational Excellence",
    "Johnson Controls releases second data center reference design guide to advance industrial‑scale AI factory cooling",
    "The $3 Billion Blind Spot: How Nevada's Cooling Tower Ban Is Shifting 100 Million Gallons of Hidden Water Consumption to Power Plants",
    "Hotter Than a Hot Tub: The 45°C Breakthrough to Cool AI’s Biggest Machines",
    "Knorr-Bremse Previews Braking, Steering Tech at IAA 2026",
}

rows = []
seen_urls = set()
for d in data:
    md = d.get("metadata", {})
    kw = md.get("keyword", "")
    if md.get("sourceType") != "keyword" or kw not in KEEP_KW:
        continue
    if not in_window(d.get("publishedAt", "")):
        continue
    title = d.get("title", "")
    if title in EXCLUDE_TITLES:
        continue
    url = d.get("url", "")
    if url in seen_urls:
        continue
    seen_urls.add(url)

    t_low = title.lower()
    found_competitor = None
    for alias, canon in ALIAS_TO_CANON.items():
        if alias in t_low:
            found_competitor = canon
            break
    is_our_brand = any(a in t_low for a in OUR_BRAND_ALIASES)

    themes = classify_themes(title)
    theme_primary = themes[0] if themes else "General brand activity"
    sentiment = classify_sentiment(title)
    meaning, team = THEME_INSIGHT.get(theme_primary, THEME_INSIGHT["General brand activity"])

    rows.append({
        "Date": d.get("publishedAt", "")[:10],
        "Title": title,
        "Source": d.get("source", ""),
        "URL": url,
        "Keyword bucket": kw,
        "Competitor": found_competitor,
        "Is Atlas Copco": is_our_brand,
        "Theme": theme_primary,
        "Sentiment": sentiment,
        "Possible meaning": meaning,
        "Team to review": team,
    })

df = pd.DataFrame(rows).sort_values("Date").reset_index(drop=True)
pd.set_option("display.max_colwidth", 60)
pd.set_option("display.width", 220)
print("Total kept rows:", len(df))
print()
print("=== Competitor-specific rows ===")
comp_rows = df[df["Competitor"].notna()]
print(comp_rows[["Date", "Competitor", "Title", "Theme", "Sentiment"]].to_string())
print()
print("=== Atlas Copco's own coverage ===")
ac_rows = df[(df["Is Atlas Copco"]) & (df["Competitor"].isna())]
print(ac_rows[["Date", "Title", "Theme", "Sentiment"]].to_string())
print()
print("=== Industry-wide (no specific company) ===")
ind_rows = df[(df["Competitor"].isna()) & (~df["Is Atlas Copco"])]
print(ind_rows[["Date", "Keyword bucket", "Title", "Theme", "Sentiment"]].to_string())
print()
print("Counts by competitor:", comp_rows["Competitor"].value_counts().to_dict())
print("Atlas Copco rows:", len(ac_rows))
print("Industry rows:", len(ind_rows))

df.to_csv("/tmp/newsetl/news_classified.csv", index=False)
