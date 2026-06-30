# -*- coding: utf-8 -*-
"""Stage 6: Opportunities, Raw Data Explorer, Executive Summary, Settings."""
import pandas as pd
import pickle

with open("sheets_stage1.pkl", "rb") as f:
    df, SHEETS = pickle.load(f)

TRACKED_COMPETITORS = [
    "Ingersoll Rand", "Sullair", "ELGi", "FS-Curtis", "Quincy Compressor",
    "Kaeser Compressors", "BOGE", "Hitachi Global Air Power", "Gardner Denver",
]
OUR_BRAND = "Atlas Copco"

# ---------------------------------------------------------------------------
# Clean_Opportunities  (grounded in verified evidence pulled from the data;
# soft "for review" language throughout per dashboard brief)
# ---------------------------------------------------------------------------
# News-coverage opportunity (computed live from the data, not hardcoded, so it
# never goes stale if more Google News signals are added in a future period)
_news_df = df[df["channel"] == "News"]
_news_by_comp = _news_df[_news_df["competitor"].isin(TRACKED_COMPETITORS)]["competitor"].value_counts()
_covered = [c for c in TRACKED_COMPETITORS if c in _news_by_comp.index]
_uncovered = [c for c in TRACKED_COMPETITORS if c not in _news_by_comp.index]
if len(_news_df) == 0:
    _news_opp = dict(
        opp="No PR releases, news mentions, dedicated event pages, or webinar listings were confirmed for any tracked competitor this period",
        evid="0 of " + str(len(df)) + " records this period map to the PR, News, Events, or Webinars channels after data-quality review.",
        why="This may reflect genuinely limited competitor PR/event activity in May-June 2026, or it may indicate a gap in monitoring coverage for these specific channels. Recommend validating directly with the PR/Comms team before concluding no activity occurred.",
        team="PR/Comms (validate coverage gap)", priority="Medium", confidence="Low")
else:
    _covered_phrase = "; ".join(f"{c} ({_news_by_comp[c]} article{'s' if _news_by_comp[c] != 1 else ''})" for c in _covered)
    _uncovered_phrase = ", ".join(_uncovered) if _uncovered else "none"
    _news_opp = dict(
        opp="Google News coverage surfaced real PR/news-style signals for some, but not all, tracked competitors this period",
        evid=f"{len(_news_df)} Google News articles matched to the air-compressor category this period (May-June 2026). Competitor-specific coverage found for: {_covered_phrase}. "
             f"No confirmed news coverage found this period for: {_uncovered_phrase}. Several keyword-bucket matches (e.g. homonym/local-interest results for \"Gardner Denver\" and \"Quincy Compressor\" search terms) were manually excluded as false positives -- see methodology notes.",
        why="Uneven news coverage across competitors may reflect real differences in PR/media activity, or it may simply reflect what a keyword-based news scrape happens to surface. Recommend the PR/Comms team validate before drawing conclusions about competitors with no confirmed coverage.",
        team="PR/Comms (validate coverage)", priority="Medium", confidence="Medium")

OPPS = [
    dict(opp="Ingersoll Rand pairing employer-brand visibility with active hiring growth",
         evid="Employee post on 2026 Fortune 500 listing (1,831 engagements) + company post (806 engagements); "
              "workforce data shows 1,033 open roles, +38% growth over 3 months, concentrated in Engineering. "
              "https://www.linkedin.com/posts/vicente-reynal-02939b8_makinglifebetter-ugcPost-7467874064410476544-fWia",
         why="High-engagement brand-recognition content alongside accelerating hiring may indicate IR is in an active growth/expansion phase. Could be relevant context for competitive positioning discussions.",
         team="Leadership; Talent/HR; PR", priority="High", confidence="High"),
    dict(opp="ELGi concentrating visibility around INTEC 2026 trade show + executive thought leadership",
         evid="Top 3 highest-engagement Instagram posts this period are all INTEC 2026 booth/award content (378, 193, 184 engagements); "
              "MD Dr. Jairam Varadaraj featured in press interviews. https://www.instagram.com/p/DZR_-svJT8z/",
         why="Heavy in-person event + executive-visibility push may be worth reviewing for event/PR participation in the same circuits, or for content-format ideas (founder/exec-led storytelling performs well).",
         team="Events/PR; Product Marketing", priority="High", confidence="High"),
    dict(opp="Kaeser Compressors sustaining the highest LinkedIn posting cadence of all tracked competitors",
         evid="49 original LinkedIn posts in the period (highest of the 9 tracked competitors), heavily weighted toward maintenance/reliability and employee-culture content. "
              "https://www.linkedin.com/posts/kaeser-compressors_maintenance-checklist-infographic-pdfpdf-activity-7467545799162220545-sBLY",
         why="A high, steady posting cadence combined with culture-forward content may be building consistent community engagement. May be worth reviewing posting frequency/format benchmarks.",
         team="Social/Content", priority="Medium", confidence="High"),
    dict(opp="FS-Curtis using limited-edition co-branded merchandise as an engagement tactic",
         evid="\"Toledo Tools\" Limited Edition OG Yellow Line post drew 39 engagements on LinkedIn, its top post this period. "
              "https://www.linkedin.com/posts/fs-curtis_toledotools-limitededition-industrialequipment-activity-7467178201119739906-E3JD",
         why="A lower-cost merchandise/community tactic appears to outperform routine maintenance posts for this brand. Could be a low-priority idea worth a content-team look.",
         team="Social/Content", priority="Low", confidence="Medium"),
    dict(opp="Quincy Compressor posting direct hiring calls for supply-chain/procurement roles on social",
         evid="\"Procurement & Planning Specialist\" and \"Strategic Sourcing Specialist\" roles promoted directly via LinkedIn/employee posts, alongside 5 dated blog posts this period. "
              "https://www.linkedin.com/posts/sneha-kolukula_easy-apply-activity-7473857167800889344-R01y",
         why="Promoting hiring needs directly on social (rather than only job boards) may point to supply-chain/procurement build-out. Worth keeping on the radar for sales/territory context.",
         team="Talent/HR; Sales", priority="Medium", confidence="Medium"),
    dict(opp="Sullair shows limited confirmed public content this period after data-quality review",
         evid="Top content is routine maintenance messaging (54 engagements: \"Take care of your equipment...\"); workforce data shows 0 current open roles; "
              "several candidate PR/news/product-page items from the source scrape were excluded because they relied on an unreliable page-level timestamp rather than a true publish date (see Raw Data Explorer notes). "
              "https://www.linkedin.com/posts/sullair-llc_maintenance-musts-for-portable-air-compressors-activity-7467606173525884929-1CQX",
         why="Could reflect genuinely lower public-facing activity this period, or could be a monitoring/coverage gap rather than true inactivity. Recommend validating directly before drawing conclusions.",
         team="PR/Marketing (validate coverage)", priority="Medium", confidence="Low"),
    dict(opp="BOGE combining frontline-hiring posts with a value-tier product push",
         evid="\"WE ARE HIRING: 3 Field Services Engineers\" post (19 engagements) and YouTube launch of the \"BOGE E-Series\" budget compressor (17 engagements). "
              "https://www.youtube.com/watch?v=YdaIPYU1gIQ",
         why="Service-capacity hiring alongside a budget-tier product launch could point to a push into price-sensitive segments or expanded service coverage. May be worth a product-marketing review.",
         team="Product Marketing; Sales", priority="Medium", confidence="Medium"),
    dict(opp="Hitachi Global Air Power leaning into sustainability/heritage storytelling while hiring activity cools",
         evid="\"Earth Month EcoChallenge\" winner post (139 engagements) and 60th-anniversary-of-Sullair-brand heritage post (179 engagements); "
              "workforce data shows open roles down 12% over 3 months and 46% over 6 months. https://www.linkedin.com/posts/hitachi-global-air-power_our-winner-of-the-earth-month-ecochallenge-activity-7470142947523543043-vbAG",
         why="Strong sustainability/culture content paired with a cooling hiring trend may simply reflect seasonality, but the combination may be worth a periodic check-in rather than a one-time read.",
         team="Leadership; Marketing", priority="Medium", confidence="Medium"),
    dict(opp="Gardner Denver shows mostly third-party/aftermarket mentions rather than first-party brand content",
         evid="Most \"Gardner Denver\" mentions this period come from independent repair shops and parts dealers referencing GD equipment, not Gardner Denver's own channels; "
              "one confirmed first-party post found (Spanish-language Dry Cooler CDC product post). https://www.linkedin.com/posts/gardner-denver_gardnerdenver-enfriamientoindustrial-operaciaejndeplanta-activity-7477345163439329281-xZ-g",
         why="Lower first-party visibility relative to other tracked brands in this dataset could reflect genuinely lower content cadence, or a monitoring/coverage gap. Recommend validating before concluding either way.",
         team="Marketing (validate coverage)", priority="Low", confidence="Low"),
    _news_opp,
    dict(opp="\"Maintenance / repair\" is the single largest content theme and now spans all 9 tracked competitors",
         evid="62 items classified to Maintenance/repair this period, present across all 9 tracked competitors (largest theme by frequency).",
         why="This theme may now be a category-wide baseline expectation rather than a differentiator. Could be an opportunity to review whether a more distinctive angle is worth pursuing.",
         team="Content/SEO", priority="Medium", confidence="High"),
    dict(opp="Sustainability/ESG messaging is appearing across a majority of tracked competitors",
         evid="26 items classified to Sustainability this period, spanning 6 of the 9 tracked competitors.",
         why="A broadening category-wide sustainability narrative may be relevant for the content calendar and PR positioning to review.",
         team="Content/Marketing/PR", priority="Medium", confidence="High"),
    dict(opp="Multiple signals point to India/APAC market activity among competitors",
         evid="ELGi executive media interviews referencing India and Europe; Hitachi-affiliated hiring posts referencing India-based roles; geography tags found on 104 of 409 records overall.",
         why="A cluster of India/APAC-tagged activity could indicate increased competitor focus in that region. May be relevant for regional sales/marketing teams to review.",
         team="Regional Sales/Marketing", priority="Medium", confidence="Medium"),
    dict(opp="No negative-sentiment signal detected in direct mentions of our brand this period",
         evid="19 LinkedIn mentions of Atlas Copco found this period: 13 neutral, 5 positive, 0 negative, 0 mixed.",
         why="This is a useful baseline reading rather than an action item. Recommend continued routine monitoring; no response action appears warranted from this data alone.",
         team="Social/PR (monitoring only)", priority="Low", confidence="Medium"),
]

opp_df = pd.DataFrame(OPPS).rename(columns={
    "opp": "Opportunity/Signal", "evid": "Evidence", "why": "Why it may matter",
    "team": "Suggested team to review", "priority": "Priority", "confidence": "Confidence",
})
prio_rank = {"High": 0, "Medium": 1, "Low": 2}
opp_df = opp_df.sort_values(by="Priority", key=lambda s: s.map(prio_rank)).reset_index(drop=True)
SHEETS["Clean_Opportunities"] = opp_df

# ---------------------------------------------------------------------------
# Clean_Raw_Data  (full explorer)
# ---------------------------------------------------------------------------
raw = df.copy()
# Where no article body was scraped (e.g. Google News headline-only records),
# fall back to the real title rather than leaving the Summary column blank.
raw["content_short"] = raw["content"].where(raw["content"].str.len() > 0, raw["title"]).str.slice(0, 300)
raw_out = raw[[
    "date", "month", "competitor", "source_label", "channel", "title", "url", "content_short",
    "theme_primary", "product_category", "geography", "sentiment", "engagement",
    "likes", "comments", "shares", "views", "content_format",
    "possible_meaning", "team_to_review", "priority", "confidence", "mention_type",
]].rename(columns={
    "date": "Date", "month": "Month", "competitor": "Competitor", "source_label": "Source type",
    "channel": "Channel", "title": "Title", "url": "URL", "content_short": "Summary",
    "theme_primary": "Theme", "product_category": "Product/service category", "geography": "Geography",
    "sentiment": "Sentiment", "engagement": "Engagement", "likes": "Likes",
    "comments": "Comments", "shares": "Shares", "views": "Views",
    "content_format": "Content format", "possible_meaning": "Possible meaning",
    "team_to_review": "Team to review", "priority": "Priority", "confidence": "Confidence",
    "mention_type": "Mention type",
})
raw_out = raw_out.sort_values(["Month", "Date"]).reset_index(drop=True)
SHEETS["Clean_Raw_Data"] = raw_out

# ---------------------------------------------------------------------------
# Clean_Executive_Summary  (cards + narrative)
# ---------------------------------------------------------------------------
total_tracked_activity = int(df["competitor"].isin(TRACKED_COMPETITORS).sum())
total_competitors = len(TRACKED_COMPETITORS)
social_mask = df["channel"].isin(["LinkedIn", "Instagram", "YouTube"]) & (df["mention_type"] != "Comment")
total_social_posts = int(social_mask.sum())
total_blogs = int((df["channel"] == "Blog").sum())
total_product_pages = int((df["channel"] == "Product pages").sum())
total_events_webinars = int(df["channel"].isin(["Events", "Webinars"]).sum())
total_pr_news = int(df["channel"].isin(["PR", "News"]).sum())
total_jobs = int((df["channel"] == "Jobs").sum())
confirmed_webinars = int((df["channel"] == "Webinars").sum())
training_webinar_signals = int((df["theme_primary"] == "Training / webinar").sum())
confirmed_pr_releases = int((df["channel"] == "PR").sum())
news_mentions = int((df["channel"] == "News").sum())
social_engagement = int(df.loc[social_mask, "engagement"].sum())

theme_counts_excl_general = df[df["theme_primary"] != "General brand activity"]["theme_primary"].value_counts()
top3_themes = theme_counts_excl_general.head(3).index.tolist()
top3_opps = opp_df[opp_df["Priority"] == "High"]["Opportunity/Signal"].tolist()[:3]
if len(top3_opps) < 3:
    top3_opps += opp_df[opp_df["Priority"] == "Medium"]["Opportunity/Signal"].tolist()[: 3 - len(top3_opps)]

chan_counts_overall = df["channel"].value_counts()
top_channels = chan_counts_overall.head(2).index.tolist()
top_channel_labels = {"LinkedIn": "LinkedIn", "Instagram": "Instagram", "YouTube": "YouTube",
                       "Blog": "blog content", "Jobs": "hiring/workforce data", "Employee posts": "employee posts"}
channels_phrase = " and ".join(top_channel_labels.get(c, c) for c in top_channels)
themes_phrase = ", ".join(top3_themes)

narrative = (
    f"During June 2026, competitor activity was mainly concentrated around {themes_phrase}. "
    f"The strongest signals came from {channels_phrase}. "
    f"These may be relevant for content, social, product marketing, and sales teams to review."
)

exec_rows = [
    ("Total competitor activities tracked (May+June 2026)", total_tracked_activity),
    ("Total competitors tracked", total_competitors),
    ("Total social posts (LinkedIn + Instagram + YouTube)", total_social_posts),
    ("Total social engagement", social_engagement),
    ("Total blog posts", total_blogs),
    ("Total new product/service pages", total_product_pages),
    ("Total events + webinars", total_events_webinars),
    ("Confirmed webinar listings", confirmed_webinars),
    ("Training/webinar-themed signals", training_webinar_signals),
    ("Total PR releases + news mentions", total_pr_news),
    ("Confirmed PR releases", confirmed_pr_releases),
    ("News mentions", news_mentions),
    ("Total LinkedIn jobs / workforce signals", total_jobs),
    ("Top theme #1", top3_themes[0] if len(top3_themes) > 0 else "No signals found this period"),
    ("Top theme #2", top3_themes[1] if len(top3_themes) > 1 else "No signals found this period"),
    ("Top theme #3", top3_themes[2] if len(top3_themes) > 2 else "No signals found this period"),
    ("Possible opportunity #1", top3_opps[0] if len(top3_opps) > 0 else "No signals found this period"),
    ("Possible opportunity #2", top3_opps[1] if len(top3_opps) > 1 else "No signals found this period"),
    ("Possible opportunity #3", top3_opps[2] if len(top3_opps) > 2 else "No signals found this period"),
    ("Executive summary narrative", narrative),
    ("Disclaimer", "This dashboard surfaces external competitor signals for review. Final actions should be validated by the relevant marketing, product, sales, or leadership teams."),
]
SHEETS["Clean_Executive_Summary"] = pd.DataFrame(exec_rows, columns=["Metric", "Value"])

# ---------------------------------------------------------------------------
# Settings
# ---------------------------------------------------------------------------
settings_rows = [
    ("Our_Brand", OUR_BRAND),
    ("Report_Month", "June 2026"),
    ("Baseline_Month", "May 2026"),
    ("Competitors_Tracked", "; ".join(TRACKED_COMPETITORS)),
    ("Total_Records", len(df)),
    ("Generated_Date", pd.Timestamp.now().strftime("%Y-%m-%d")),
    ("Primary_Brand_Color", "#0099CC"),
    ("Brand_Color_Name", "Atlas Copco Blue (PMS 313 C / RAL 5012)"),
    ("Methodology_Note_1", "Records with date_source == 'time-tag' from the source scrape were excluded: this field was found to assign an identical page-refresh timestamp to many historically old, unrelated articles rather than a true publish date."),
    ("Methodology_Note_2", "Generic syndicated market-research-report spam posts (e.g. accounts posting boilerplate market-size/CAGR copy that namedrops multiple brands) were excluded as non-signals."),
    ("Methodology_Note_3", "Categories with zero confirmed records (e.g. PR, News, Events, Webinars in this period) are shown as 'No signals found this period' rather than estimated or invented."),
    ("Methodology_Note_4", "Google News results were matched to the air-compressor category by keyword search. Results that were clearly homonym/false-positive matches (e.g. Denver, Colorado garden articles matched on \"Gardner Denver\"; Quincy, Illinois local-interest articles matched on \"Quincy Compressor\"; hair-dryer reviews matched on \"air dryer\") were manually reviewed and excluded as non-signals rather than included as-is."),
    ("Disclaimer", "This dashboard surfaces external competitor signals for review. Final actions should be validated by the relevant marketing, product, sales, or leadership teams."),
]
SHEETS["Settings"] = pd.DataFrame(settings_rows, columns=["Setting", "Value"])

print("All sheets ready:", list(SHEETS.keys()))
for k, v in SHEETS.items():
    print(f"  {k}: {v.shape}")

with open("sheets_final.pkl", "wb") as f:
    pickle.dump(SHEETS, f)
