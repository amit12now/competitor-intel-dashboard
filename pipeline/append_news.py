# -*- coding: utf-8 -*-
"""
Append real Google News signals (industry + Atlas Copco + competitors) into
the stage3_scoped raw-record table as channel="News" rows, so the existing
enrich.py -> build_master_excel.py -> build_master_excel_stage2.py ->
write_master_xlsx.py pipeline classifies/aggregates them exactly the same
way it does every other record (no hand-fabricated fields).
"""
import pandas as pd

news = pd.read_csv("news_classified.csv", keep_default_na=False)
stage3 = pd.read_csv("stage3_scoped.csv", keep_default_na=False)

# next record_id
existing_nums = stage3["record_id"].str.replace("R", "", regex=False).astype(int)
next_num = existing_nums.max() + 1

rows = []
for i, r in news.iterrows():
    comp = r["Competitor"].strip()
    is_ac = str(r["Is Atlas Copco"]).strip().lower() == "true"

    if comp:
        competitor = comp
        all_companies = comp
        mention_type = "Competitor Mention"
    elif is_ac:
        competitor = "Atlas Copco"
        all_companies = "Our Brand"
        mention_type = "Our Brand Mention"
    else:
        competitor = "Unspecified / General"
        all_companies = ""
        mention_type = "Air Compressor Keyword"

    date_str = r["Date"] + "T08:00:00.000Z"  # only a date was scraped; noon-ish placeholder time, real date
    month = "May 2026" if date_str[:7] == "2026-05" else "June 2026"

    rows.append({
        "record_id": f"R{next_num + i:05d}",
        "date": date_str,
        "month": month,
        "competitor": competitor,
        "all_companies_mentioned": all_companies,
        "source_label": "Google News article",
        "channel": "News",
        "mention_type": mention_type,
        "title": r["Title"],
        "content": "",
        "url": r["URL"],
        "author_name": r["Source"],
        "author_type": "media",
        "engagement": 0,
        "job_location": "",
        "job_top_function": "",
    })

new_df = pd.DataFrame(rows)
assert list(new_df.columns) == list(stage3.columns), (list(new_df.columns), list(stage3.columns))

combined = pd.concat([stage3, new_df], ignore_index=True)
combined.to_csv("stage3_scoped.csv", index=False)
print("Appended", len(new_df), "rows. New stage3_scoped.csv shape:", combined.shape)
print(new_df["mention_type"].value_counts())
print(new_df["competitor"].value_counts())
