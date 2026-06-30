# -*- coding: utf-8 -*-
"""
Loads every raw scraped source file in indata/ and normalizes into a single
list of record dicts following a unified schema. No fabrication: if a field
can't be determined from the source, it is left blank.
"""
import json
import re
import pandas as pd
from datetime import datetime
from taxonomy import (
    find_companies_in_text, mentions_our_brand, mentions_industry_keyword,
    OUR_BRAND, COMPETITOR_LIST,
)

IN = ".."

def load_json(name):
    with open(f"{IN}/{name}", "r", encoding="utf-8") as fp:
        return json.load(fp)

def month_label(date_iso):
    if not date_iso:
        return ""
    try:
        d = datetime.fromisoformat(date_iso.replace("Z", "+00:00"))
    except Exception:
        return ""
    if d.year == 2026 and d.month == 6:
        return "June 2026"
    if d.year == 2026 and d.month == 5:
        return "May 2026"
    return f"{d.strftime('%B %Y')}"

def engagement_sum(eng):
    if not isinstance(eng, dict):
        return 0
    return int(eng.get("likes") or 0) + int(eng.get("comments") or 0) + int(eng.get("shares") or 0)

RECORDS = []
_rid = [0]
def new_id():
    _rid[0] += 1
    return f"R{_rid[0]:05d}"

def add_record(**kw):
    rec = {
        "record_id": new_id(),
        "date": "", "month": "", "competitor": "", "all_companies_mentioned": "",
        "source_label": "", "channel": "", "mention_type": "", "title": "",
        "content": "", "url": "", "author_name": "", "author_type": "",
        "engagement": 0,
    }
    rec.update(kw)
    RECORDS.append(rec)

# ---------------------------------------------------------------------------
# 1) LinkedIn JSON files (post / comment / reaction)
# ---------------------------------------------------------------------------
LINKEDIN_FILES = [
    ("Linkedin-company-posts.json", "LinkedIn company posts"),
    ("20 Posts with Competitors' Name.json", "LinkedIn posts mentioning competitor names"),
    ("20 Posts with Competitors-2 Name.json", "LinkedIn posts mentioning competitor names"),
    ("Company Thought Posts.json", "LinkedIn posts mentioning 'air compressor'"),
    ("Indusrtrial air compressor -linkedin.json", "LinkedIn posts mentioning 'air compressor'"),
    ("atlas copco.json", "LinkedIn posts mentioning our brand"),
]

def process_linkedin_file(fname, default_label):
    data = load_json(fname)
    posts = [d for d in data if d.get("type") == "post"]
    comments = [d for d in data if d.get("type") == "comment"]

    # map postId -> company canonical name if the post itself resolves to a tracked company author
    postid_company = {}
    for p in posts:
        author = p.get("author", {}) or {}
        if author.get("type") == "company":
            companies = find_companies_in_text(author.get("name", ""))
            if companies:
                postid_company[p.get("id")] = sorted(companies)[0]

    kept, dropped = 0, 0
    for p in posts:
        content = p.get("content") or ""
        author = p.get("author", {}) or {}
        author_name = author.get("name", "") or ""
        author_type = author.get("type", "") or ""
        author_info = author.get("info", "") or ""
        full_text = " ".join([content, author_name, author_info])

        companies_in_author = find_companies_in_text(author_name)
        companies_in_text = find_companies_in_text(full_text)
        our_brand = mentions_our_brand(full_text)
        industry_kw = mentions_industry_keyword(full_text)

        if not companies_in_text and not our_brand and not industry_kw:
            dropped += 1
            continue  # irrelevant noise (e.g. unrelated local-business posts)

        date_iso = (p.get("postedAt") or {}).get("date", "")
        url = p.get("linkedinUrl") or p.get("shareLinkedinUrl") or ""
        eng = engagement_sum(p.get("engagement"))

        # Determine attribution
        if author_type == "company" and companies_in_author:
            competitor = sorted(companies_in_author)[0]
            mention_type = "Company Post"
            channel = "LinkedIn"
            source_label = "LinkedIn company posts"
        elif author_type == "profile" and (companies_in_author or any(c.lower() in author_info.lower() for c in COMPETITOR_LIST)):
            competitor = sorted(companies_in_author)[0] if companies_in_author else (sorted(companies_in_text)[0] if companies_in_text else "")
            mention_type = "Employee Post"
            channel = "Employee posts"
            source_label = "LinkedIn posts from company employees"
        elif our_brand and not companies_in_text:
            competitor = OUR_BRAND
            mention_type = "Our Brand Mention"
            channel = "LinkedIn"
            source_label = "LinkedIn posts mentioning our brand"
        elif companies_in_text:
            competitor = sorted(companies_in_text)[0]
            mention_type = "Competitor Mention"
            channel = "LinkedIn"
            source_label = default_label if "competitor" in default_label else "LinkedIn posts mentioning competitor names"
        else:
            competitor = "Unspecified / General"
            mention_type = "Air Compressor Keyword"
            channel = "LinkedIn"
            source_label = "LinkedIn posts mentioning 'air compressor'"

        kept += 1
        add_record(
            date=date_iso, month=month_label(date_iso), competitor=competitor,
            all_companies_mentioned="; ".join(sorted(companies_in_text)) if companies_in_text else ("Our Brand" if our_brand else ""),
            source_label=source_label, channel=channel, mention_type=mention_type,
            title=content[:90].replace("\n", " ").strip(),
            content=content.strip(), url=url, author_name=author_name,
            author_type=author_type, engagement=eng,
        )

    for c in comments:
        commentary = c.get("commentary") or ""
        actor = c.get("actor", {}) or {}
        actor_name = actor.get("name", "") or ""
        actor_position = actor.get("position", "") or ""
        full_text = " ".join([commentary, actor_position])
        companies_in_text = find_companies_in_text(full_text)
        our_brand = mentions_our_brand(full_text)
        industry_kw = mentions_industry_keyword(full_text)

        post_company = postid_company.get(c.get("postId") or (c.get("query") or {}).get("post"))
        if not companies_in_text and not our_brand and not industry_kw and not post_company:
            continue

        competitor = post_company or (sorted(companies_in_text)[0] if companies_in_text else (OUR_BRAND if our_brand else "Unspecified / General"))
        date_iso = c.get("createdAt") or ""
        eng = engagement_sum(c.get("engagement"))
        add_record(
            date=date_iso, month=month_label(date_iso), competitor=competitor,
            all_companies_mentioned="; ".join(sorted(companies_in_text)) if companies_in_text else "",
            source_label="LinkedIn comments (market reaction)", channel="LinkedIn",
            mention_type="Comment",
            title=commentary[:90].replace("\n", " ").strip(),
            content=commentary.strip(), url=c.get("linkedinUrl", ""),
            author_name=actor_name, author_type="profile", engagement=eng,
        )
    return kept, dropped

for fname, label in LINKEDIN_FILES:
    k, d = process_linkedin_file(fname, label)
    print(f"{fname}: kept {k}, dropped {d} (noise)")

print("Total records after LinkedIn pass:", len(RECORDS))

import pandas as pd
df_check = pd.DataFrame(RECORDS)
df_check.to_csv("stage1_linkedin.csv", index=False)
print(df_check["competitor"].value_counts())
print(df_check["channel"].value_counts())
print(df_check["mention_type"].value_counts())

# ---------------------------------------------------------------------------
# 2) Instagram
# ---------------------------------------------------------------------------
IG_OWNER_MAP = {
    "kaeserusa": "Kaeser Compressors",
    "ingersollrand": "Ingersoll Rand",
    "elgiaircompressors": "ELGi",
    "fs_curtis": "FS-Curtis",
    "bogekompressoren": "BOGE",
}

def process_instagram():
    data = load_json("Instgram Posts.json")
    kept, dropped = 0, 0
    for d in data:
        if d.get("error"):
            continue
        owner = d.get("ownerUsername", "")
        competitor = IG_OWNER_MAP.get(owner)
        caption = d.get("caption") or ""
        if not competitor:
            # fall back to text-based matching for unmapped accounts
            companies = find_companies_in_text(caption)
            if companies:
                competitor = sorted(companies)[0]
            elif mentions_industry_keyword(caption):
                competitor = "Unspecified / General"
            else:
                dropped += 1
                continue
        ts = d.get("timestamp", "")
        eng = int(d.get("likesCount") or 0) + int(d.get("commentsCount") or 0)
        add_record(
            date=ts, month=month_label(ts), competitor=competitor,
            all_companies_mentioned=competitor,
            source_label="Instagram posts", channel="Instagram", mention_type="Company Post",
            title=caption[:90].replace("\n", " ").strip(),
            content=caption.strip(), url=d.get("url", ""),
            author_name=d.get("ownerFullName", owner), author_type="company", engagement=eng,
        )
        kept += 1
    print(f"Instagram: kept {kept}, dropped {dropped}")

process_instagram()

# ---------------------------------------------------------------------------
# 3) YouTube
# ---------------------------------------------------------------------------
YT_CHANNEL_MAP = {
    "ELGi Air Compressors": "ELGi",
    "Kaeser USA": "Kaeser Compressors",
    "BOGE EN Compressed Air & Gas Solutions": "BOGE",
    "Ingersoll Rand Air Compressors": "Ingersoll Rand",
    "Sullair": "Sullair",
    "FS-Curtis": "FS-Curtis",
}

def process_youtube():
    data = load_json("Youtube Posts.json")
    kept, dropped = 0, 0
    for d in data:
        if d.get("error"):
            dropped += 1
            continue
        ch = d.get("channelName", "")
        competitor = YT_CHANNEL_MAP.get(ch, "")
        title = d.get("title") or ""
        if not competitor:
            companies = find_companies_in_text(title + " " + ch)
            competitor = sorted(companies)[0] if companies else "Unspecified / General"
        date_iso = d.get("date", "")
        eng = int(d.get("viewCount") or 0) // 10 + int(d.get("likes") or 0)  # downweight raw views vs likes
        add_record(
            date=date_iso, month=month_label(date_iso), competitor=competitor,
            all_companies_mentioned=competitor,
            source_label="YouTube videos", channel="YouTube", mention_type="Company Post",
            title=title[:90], content=(d.get("channelDescription") or "")[:400],
            url=d.get("url", ""), author_name=ch, author_type="company", engagement=eng,
        )
        kept += 1
    print(f"YouTube: kept {kept}, dropped {dropped}")

process_youtube()

# ---------------------------------------------------------------------------
# 4) Blog / Company News / Press Coverage / Product News / Events (debug xlsx)
# ---------------------------------------------------------------------------
CATEGORY_TO_CHANNEL = {
    "Blog": "Blog",
    "Company News": "PR",
    "Press Coverage": "News",
    "Product News": "Product pages",
    "Press / Product Launches": "Product pages",
    "Events": "Events",
}

def process_debug_excel():
    df = pd.read_excel(f"{IN}/checked_urls_debug_2026_6.xlsx")
    kept, dropped = 0, 0
    for _, row in df.iterrows():
        year = row.get("published_year")
        mon = row.get("published_month")
        status = row.get("status")
        date_source = str(row.get("date_source", ""))
        if status != "Date found":
            dropped += 1
            continue
        if date_source == "time-tag":
            # time-tag dates proved unreliable on QA (page-level "last checked"
            # indicator, not a per-article publish date -- it assigned the same
            # single date to 30 unrelated historical award announcements).
            # Excluded rather than risk presenting stale content as new signals.
            dropped += 1
            continue
        try:
            year = int(year); mon = int(mon)
        except Exception:
            dropped += 1
            continue
        if year != 2026 or mon not in (5, 6):
            dropped += 1
            continue
        date_iso = str(row.get("published_date"))
        mlabel = "June 2026" if mon == 6 else "May 2026"
        competitor = row.get("company", "Unspecified / General")
        cat = row.get("source_category", "")
        channel = CATEGORY_TO_CHANNEL.get(cat, "Blog")
        title = str(row.get("title", ""))
        add_record(
            date=date_iso, month=mlabel, competitor=competitor,
            all_companies_mentioned=competitor,
            source_label=f"{cat} pages/articles", channel=channel, mention_type="Company Post",
            title=title[:120], content=title, url=row.get("url", ""),
            author_name=competitor, author_type="company", engagement=0,
        )
        kept += 1
    print(f"Debug excel (blog/news/events/product): kept {kept}, dropped {dropped}")

process_debug_excel()

print("=" * 60)
print("TOTAL RECORDS:", len(RECORDS))
df_all = pd.DataFrame(RECORDS)
df_all.to_csv("stage2_all_raw.csv", index=False)
print(df_all["channel"].value_counts())
print()
print(df_all["competitor"].value_counts())
print()
print(df_all["month"].value_counts())

# ---------------------------------------------------------------------------
# 5) Jobs (company-level aggregate workforce analytics -> hiring signal records)
# ---------------------------------------------------------------------------
JOB_COMPANY_MAP = {
    "Ingersoll Rand": "Ingersoll Rand",
    "Ingersoll Rand Compressor Systems & Services": "Ingersoll Rand",
    "Kaeser Compressors USA": "Kaeser Compressors",
    "Hitachi Global Air Power": "Hitachi Global Air Power",
    "Sullair": "Sullair",
    "ELGi North America": "ELGi",
    "FS-Curtis": "FS-Curtis",
}

def _val_for_month(series, ym):
    for item in (series or []):
        if item.get("date") == ym:
            return item
    return None

def process_jobs():
    data = load_json("Company Job data.json")
    kept = 0
    for c in data:
        raw_name = c.get("company_name", "")
        competitor = JOB_COMPANY_MAP.get(raw_name, raw_name)
        june = _val_for_month(c.get("new_hires"), "2026-6")
        may = _val_for_month(c.get("new_hires"), "2026-5")
        total_open = c.get("total_job_openings")
        growth = c.get("job_openings_growth") or {}
        hq = ", ".join([x for x in [c.get("hq_city"), c.get("hq_country")] if x]).strip(", ")

        parts = []
        if june is not None:
            parts.append(f"{june.get('total_hires', 0)} new hires recorded in June 2026")
            if june.get("senior_hires"):
                parts.append(f"({june.get('senior_hires')} senior-level)")
        if may is not None:
            parts.append(f"vs {may.get('total_hires', 0)} in May 2026")
        if total_open is not None:
            parts.append(f"Open roles (current snapshot): {total_open}")
        if growth.get("3m"):
            parts.append(f"Open-role growth last 3 months: {growth.get('3m')}")
        if growth.get("6m"):
            parts.append(f"6 months: {growth.get('6m')}")
        if hq:
            parts.append(f"HQ: {hq}")
        title = f"Hiring signal — {competitor}: " + "; ".join(parts[:2])
        content = ". ".join(parts) + "."

        if june is None and may is None and total_open in (None, 0):
            continue  # nothing meaningful to report

        top_func = ""
        of = c.get("job_openings_by_function") or []
        if of:
            latest = of[0]
            cbf = latest.get("count_by_function") or {}
            if cbf:
                top_func_name = max(cbf.items(), key=lambda kv: kv[1].get("count", 0))[0]
                top_func = f"{top_func_name} ({cbf[top_func_name].get('percentage')}% of open roles, {latest.get('date')} snapshot)"

        add_record(
            date="2026-06-30", month="June 2026", competitor=competitor,
            all_companies_mentioned=competitor,
            source_label="LinkedIn jobs (company workforce analytics)", channel="Jobs",
            mention_type="Hiring Signal", title=title[:160], content=content,
            url=c.get("linkedin_url", ""), author_name=competitor, author_type="company",
            engagement=0, job_location=hq, job_top_function=top_func,
        )
        kept += 1
    print(f"Jobs: kept {kept} hiring-signal records")

process_jobs()

# ---------------------------------------------------------------------------
# Final scope filter: keep only May/June 2026 records, drop the rest
# ---------------------------------------------------------------------------
before = len(RECORDS)
# Drop generic syndicated "market research report" spam posts (e.g. "M&E Market
# Reports", "EnergyAxis Research") that namedrop multiple brand names in passing
# inside boilerplate CAGR/market-size copy. These are not genuine competitor
# signals and would otherwise get misclassified (e.g. as Hiring/expansion).
_SPAM_KW = ["market report", "cagr", "request for sample", "billion by 20",
            "million by 20", "research report", "market worth", "market size"]
def _is_spam(rec):
    txt = (str(rec.get("title", "")) + " " + str(rec.get("content", ""))).lower()
    return any(k in txt for k in _SPAM_KW)

_before_spam = len(RECORDS)
RECORDS[:] = [r for r in RECORDS if not _is_spam(r)]
print(f"Dropped {_before_spam - len(RECORDS)} syndicated market-research spam rows")

RECORDS[:] = [r for r in RECORDS if r["month"] in ("May 2026", "June 2026")]
print(f"Scope filter: {before} -> {len(RECORDS)} (May/June 2026 only)")

df_final = pd.DataFrame(RECORDS)
df_final.to_csv("stage3_scoped.csv", index=False)
print("=" * 60)
print("FINAL STAGE-3 TOTAL:", len(df_final))
print(df_final["channel"].value_counts())
print()
print(df_final["competitor"].value_counts())
print()
print(df_final["month"].value_counts())
