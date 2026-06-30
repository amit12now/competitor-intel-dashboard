# Handoff: Atlas Copco Competitive Intelligence Dashboard

Context dump for continuing this project in a different coding agent (Codex or similar). Written 2026-06-30. Read this whole file before changing anything — most of the hard rules below exist because they were violated once and the user pushed back hard.

## 1. What this project is

A Streamlit dashboard (`dashboard.py`) for Atlas Copco that tracks 9 named air-compressor competitors — BOGE, ELGi, FS-Curtis, Gardner Denver, Hitachi Global Air Power, Ingersoll Rand, Kaeser Compressors, Quincy Compressor, Sullair — using scraped LinkedIn/Instagram/YouTube/job-postings/blog/Google-News data. Current period is June 2026 vs a May 2026 baseline. It reads exclusively from one workbook, `Competitor_Intelligence_Master.xlsx`, which is built by a 7-stage ETL pipeline (now in `pipeline/`) from raw scraped JSON/CSV/XLSX files sitting in the project root.

The audience is non-technical (marketing/product/sales/leadership), so the whole dashboard is written as "signals for review," not conclusions.

Run it with:
```
streamlit run dashboard.py
```

## 2. Hard constraints — never violate these

1. **Disclaimer text is verbatim and must never be reworded.** It lives in the `DISCLAIMER` constant near the top of `dashboard.py`:
   > "This dashboard surfaces external competitor signals for review. Final actions should be validated by the relevant marketing, product, sales, or leadership teams."
2. **Soft language only.** "signals for review," "may be worth reviewing," "could indicate," "possible opportunity" — never "final decision," "you must," or similarly directive phrasing.
3. **No fabrication, ever.** If a competitor/category has zero confirmed signals, show "No signals found this period." Never estimate, infer, or invent a number/fact that isn't traceable to a real row in the master Excel. This is the single most important rule in the project — see section 6 for the exact escalation path when data looks missing.
4. **Atlas Copco brand palette is fixed.** Don't change the hex constants at the top of `dashboard.py` (`ATLAS_BLUE = "#0099CC"`, etc. — full list in section 4).
5. **Ingersoll Rand is permanently pinned as the first competitor tab.** `PINNED_FIRST_COMPETITOR = "Ingersoll Rand"` (line ~561). The other 3 "top" tab slots are ranked live from `Total activity` each run; don't hardcode them.
6. **Don't reflexively add charts/sections.** Anything new must be built strictly from real columns already in the master Excel (or new columns added via the ETL pipeline from real source data — never invented in dashboard.py).

## 3. If data looks missing — the standing rule

If a metric/category/competitor seems to have no data in the master Excel, the rule is: **check the raw JSON source files first.** If the data exists there but didn't make it into the Excel, fix it by extending the ETL pipeline (`pipeline/*.py`) to fold it in, regenerate `Competitor_Intelligence_Master.xlsx`, and only then surface it in `dashboard.py`. Never patch dashboard.py to show a number that isn't backed by a real Excel row. This is exactly how Google News coverage was added in a prior session (see `pipeline/process_news.py` + `pipeline/append_news.py`).

## 4. dashboard.py architecture

~1484 lines, single file, reads only `Clean_*` sheets + `Settings` from the master Excel (`DEFAULT_PATH = Path(__file__).parent / "Competitor_Intelligence_Master.xlsx"`, can also upload a different copy via the sidebar).

**Color constants** (top of file):
```python
ATLAS_BLUE = "#0099CC"; ATLAS_BLUE_DARK = "#006C92"; ATLAS_BLUE_LIGHT = "#E6F5FA"
NAVY = "#1B1F2A"; GRAY = "#6B7280"; GRAY_LIGHT = "#F4F6F8"; WHITE = "#FFFFFF"
ACCENT_AMBER = "#E8A33D"; ACCENT_AMBER_LIGHT = "#FFF6E8"
ACCENT_RED = "#D8554A"; ACCENT_GREEN = "#3FA66A"; NEUTRAL_BAR = "#C7CDD4"
CATEGORICAL_PALETTE = [ATLAS_BLUE, ACCENT_AMBER, ACCENT_GREEN, ACCENT_RED, "#7C5CBF", "#3D8C97", "#C97FB0", "#5B7FA6", "#A3A86C", "#8D6E63", "#4F6D7A"]
PRIORITY_COLORS = {"High": ACCENT_RED, "Medium": ACCENT_AMBER, "Low": "#9AA5B1"}
CONFIDENCE_COLORS = {"High": ACCENT_GREEN, "Medium": ACCENT_AMBER, "Low": "#9AA5B1"}
SENTIMENT_COLORS = {"Positive": ACCENT_GREEN, "Neutral": "#9AA5B1", "Mixed": ACCENT_AMBER, "Negative": ACCENT_RED}
ACTIVITY_LEVEL_COLORS = {"High": ATLAS_BLUE, "Medium": ACCENT_AMBER, "Low": NEUTRAL_BAR}
```
`PRIORITY_ICONS` (emoji 🔴🟡🟢) still exists but is **no longer used anywhere** as of this session's redesign — left defined but dead, harmless.

**Key helpers:**
- `insight(text)` — renders a `.insight-box` with a "💡 What this means:" prefix.
- `action(text)` — renders an `.action-box` with "✅ Worth considering:".
- `mini_stat(col, label, value)` — small KPI stat block.
- `chip(text, bg, fg=WHITE)` — solid-filled pill badge. Still used for theme/sentiment/format/activity-level/channel tags elsewhere in the file. **Do not reuse this for priority/confidence badges** — that pattern was deliberately removed (see section 5).
- `info_if_empty(df, label)` — the standard no-fabrication-safe empty-state renderer ("No signals found this period").
- `_pct(n, d)` — safe percentage helper.
- `exec_metric(key)` — looks up a Metric/Value pair from the `Clean_Executive_Summary` sheet.
- `logo_badge(name, logo_map, size)` — renders a competitor logo if available in `competitor_logos.json`, else a plain letter badge (no fabricated logo).
- `style_layout(fig, height, legend)` — shared Plotly chart theming.

**CSS:** one big `<style>` block. Newest classes (added this session, sit right after `.action-why`):
```css
.card-meta { display: flex; align-items: center; flex-wrap: wrap; gap: 6px; color: {GRAY}; font-size: 0.78rem; font-weight: 600; margin-top: 8px; }
.meta-dot { width: 7px; height: 7px; border-radius: 50%; display: inline-block; flex: 0 0 auto; }
.team-tag { display: inline-block; font-size: 0.72rem; font-weight: 600; padding: 2px 9px; border-radius: 6px; background: {GRAY_LIGHT}; color: {NAVY}; border: 1px solid #E2E8F0; }
```
Current convention for priority/confidence anywhere in the Executive Summary or competitor cards: a small flat `<span class='meta-dot' style='background:{color}'></span>` + plain muted text ("High priority · High confidence"), **not** emoji + solid chip pills. Team assignments use `.team-tag` (light outline badge), not the solid `chip()` pill.

**Tab structure** (`_tab_labels`, built dynamically, line ~1144):
1. "1. Executive Summary"
2. One tab per `TOP_COMPETITORS` entry (Ingersoll Rand pinned first, then top-3 by `Total activity`) — "2. Ingersoll Rand", "3. ...", "4. ...", "5. ..."
3. "{n}. Other Competitors" — summary view for everyone not in the top 4
4. "{n+1}. Raw Data Explorer" — filterable raw table

Each competitor profile tab internally has its own `st.tabs([...])` for platform breakdowns (LinkedIn / Instagram / YouTube / Other & mentions). **Gotcha:** Streamlit's `AppTest.tabs` returns a single flattened list of every `st.tabs()` group in the entire script (currently 23 entries total across the whole file) — when testing, locate target tabs by their label/content, never by a fixed index.

## 5. What changed in the most recent session (for context, not action needed)

Two rounds of user-driven redesign, both already shipped and committed:
1. Rewrote the "Industry news pulse" insight box in the Executive Summary so it does volume-vs-substance analysis (e.g. "X led News volume but only Y% was a genuine launch/business move, vs Atlas Copco's smaller-but-more-substantive coverage") instead of just restating article counts per competitor.
2. Redesigned "Key themes to watch," "⭐ Top opportunities for review," "✅ What to do next," and the matching per-competitor "Opportunities for review" cards to remove the emoji-circle (🔴🟡🟢) + solid double chip-pill badge look, replacing it with the flat `meta-dot` + plain-text + `team-tag` pattern described above.

Both changes are in commit `cb9c4c4`.

## 6. ETL pipeline (now in `pipeline/`)

Just copied into the project (previously only existed in a throwaway sandbox mirror) and **verified end-to-end in a clean test directory** — running all 7 scripts in order from scratch reproduces a `Competitor_Intelligence_Master.xlsx` that is structurally identical to the one currently in the project root (same sheet shapes, e.g. `Clean_PR_News_Events: (47, 10)`), and the resulting file loads in `dashboard.py` with zero `AppTest` exceptions.

**Run order** (run from inside `pipeline/`, working directory matters — relative paths assume cwd = `pipeline/`):
```
cd pipeline
python3 load_sources.py              # raw JSON/XLSX -> stage1_linkedin.csv, stage2_all_raw.csv, stage3_scoped.csv
python3 append_news.py                # OPTIONAL: folds pipeline/news_classified.csv into stage3_scoped.csv as channel="News" rows
python3 enrich.py                     # stage3_scoped.csv -> stage4_enriched.csv (theme/sentiment/geography/priority/confidence)
python3 build_master_excel.py         # stage4_enriched.csv -> sheets_stage1.pkl
python3 build_master_excel_stage2.py  # sheets_stage1.pkl -> sheets_final.pkl (adds Opportunities, Raw Data, Exec Summary, Settings)
python3 write_master_xlsx.py          # sheets_final.pkl -> ../Competitor_Intelligence_Master.xlsx (writes to project root)
```

**⚠️ Important ordering gotcha:** `load_sources.py` regenerates `stage3_scoped.csv` from scratch each time, which would **drop the Google News rows** unless you re-run `append_news.py` immediately afterward (before `enrich.py`). The `Competitor_Intelligence_Master.xlsx` currently shipping already has the Google News rows folded in. If you re-run the pipeline and skip `append_news.py`, you will silently lose that data — the file will still look valid, just thinner. This was confirmed during this session's reproducibility test (with `append_news.py` included, `Clean_PR_News_Events` came out as 47 rows, matching production).

**Script-by-script summary:**
| Script | Reads | Writes | Purpose |
|---|---|---|---|
| `load_sources.py` | raw JSON/XLSX in project root (`IN = ".."`) | `stage1_linkedin.csv`, `stage2_all_raw.csv`, `stage3_scoped.csv` | Normalizes every raw scraped source into one unified record schema. Filters out syndicated "market research report" spam and scopes to May/June 2026 only. |
| `taxonomy.py` | — (shared module, no I/O) | — | Competitor alias list (`COMPETITORS`, 9 keys, exact order matches the hard-required list), `OUR_BRAND = "Atlas Copco"` + aliases, theme/sentiment/geography keyword dictionaries, `THEME_INSIGHT` template map. |
| `enrich.py` | `stage3_scoped.csv` | `stage4_enriched.csv` | Rule-based (no LLM, no fabrication) tagging: theme, sentiment, geography, priority, confidence — all derived from actual record text/metadata. |
| `build_master_excel.py` | `stage4_enriched.csv` | `sheets_stage1.pkl` | Builds the first batch of `Clean_*` sheet DataFrames. |
| `build_master_excel_stage2.py` | `sheets_stage1.pkl` | `sheets_final.pkl` | Adds Opportunities, Raw Data Explorer, Executive Summary, Settings sheets. |
| `write_master_xlsx.py` | `sheets_final.pkl` | `../Competitor_Intelligence_Master.xlsx` | Formats and saves the final workbook with Atlas Copco branding (openpyxl, header fills, column widths, `Table` objects). |
| `process_news.py` | `pipeline/dataset_google-news-scraper-fast_2026-06-30_04-31-13-598.json` (raw Google News scrape export) | `pipeline/news_classified.csv` | One-off classification of the raw Google News scrape into per-competitor/Atlas-Copco/general buckets with a keyword allowlist and a date-window filter. Kept as a worked example for the next time a fresh news scrape needs to be classified — not meant to run unattended on new data without checking the keyword/window logic still fits. |
| `append_news.py` | `pipeline/news_classified.csv`, `stage3_scoped.csv` | `stage3_scoped.csv` (overwritten in place) | Appends the classified news rows into the raw record table as channel="News" so the rest of the pipeline (enrich → build → write) treats them identically to every other record. Computes the next `record_id` automatically. |

All 7 pipeline scripts + the news-classification data (`dataset_google-news-scraper-fast_2026-06-30_04-31-13-598.json`, `news_classified.csv`) are now committed to git inside `pipeline/`.

## 7. Master Excel sheet inventory

`Competitor_Intelligence_Master.xlsx` currently has 19 sheets: `Settings`, `Clean_Executive_Summary`, `Clean_Competitor_Activity`, `Clean_Channel_Activity`, `Clean_Competitor_x_Channel`, `Clean_Trend_May_vs_June`, `Clean_Theme_Frequency`, `Clean_Theme_by_Competitor`, `Clean_Theme_by_Channel`, `Clean_Theme_Trend_May_vs_June`, `Clean_Product_Service_Signals`, `Clean_Social_Conversation`, `Clean_Conversation_Top_Themes`, `Clean_Conversation_Sentiment`, `Clean_Most_Mentioned_Brands`, `Clean_PR_News_Events`, `Clean_Hiring_Expansion`, `Clean_Opportunities`, `Clean_Raw_Data`.

## 8. Other files in the project root

- `dashboard.py` — the dashboard, tracked in git.
- `Competitor_Intelligence_Master.xlsx` — tracked in git.
- `competitor_logos.json`, `post_images.json` — small derived lookup files (logo image paths / LinkedIn post image paths keyed by competitor), tracked in git. Built from the scraped source data; if an entry is missing for a competitor, dashboard.py falls back to a plain letter badge / text card rather than inventing an image.
- `requirements.txt` — `streamlit>=1.38`, `pandas>=2.0`, `plotly>=5.20`, `openpyxl>=3.1`, `jinja2>=3.1`.
- Raw scraped source files (now tracked in git as of this session): `Linkedin-company-posts.json`, `20 Posts with Competitors' Name.json`, `20 Posts with Competitors-2 Name.json`, `Company Thought Posts.json`, `Indusrtrial air compressor -linkedin.json` (note: typo in the actual filename, not a transcription error), `atlas copco.json`, `Instgram Posts.json`, `Youtube Posts.json`, `Company Job data.json`, `checked_urls_debug_2026_6.xlsx`.
- `competitor_posts_june_2026.csv` / `.xlsx` — a blog-post scrape export (ELGi etc.). Not directly read by `load_sources.py`; looks like an earlier-stage artifact feeding into `checked_urls_debug_2026_6.xlsx` (which IS read by the pipeline). Kept for provenance.
- `pipeline/` — the ETL scripts, new this session (see section 6).
- `HANDOFF.md` — this file.

## 9. Validation pattern used throughout this project

Before considering any `dashboard.py` change safe:
```bash
python3 -m py_compile dashboard.py
python3 - << 'EOF'
from streamlit.testing.v1 import AppTest
at = AppTest.from_file("dashboard.py")
at.run(timeout=60)
print("exceptions:", at.exception)   # must be empty ElementList()
print("tabs:", len(at.tabs))         # currently 23 (flattened across all st.tabs() groups)
EOF
```
For large/risky text edits to `dashboard.py`, avoid line-based editors that can silently truncate; instead use Python `str.replace()` against a unique anchor string with an `assert content.count(anchor) == 1` guard before writing, so the script raises before any partial write. Always diff/compile/AppTest after.

## 10. Git state

Remote: `https://github.com/amit12now/competitor-intel-dashboard.git`, branch `main`. Latest commits:
```
5734686 Add ETL pipeline scripts (pipeline/) and raw source data for full reproducibility
cb9c4c4 Replace data-recap insight with volume-vs-substance analysis; redesign opportunity/action cards to remove emoji+badge clutter
b6655d9 Fix undefined chart title; add Google News industry/competitor coverage to dashboard and master Excel
```
Everything is committed locally. **Not yet pushed to origin** — this sandbox has no GitHub credentials. Push from your own machine:
```
cd "C:\Users\amith\Desktop\UBRIK\Competitor Analysis Project"
git push origin main
```

## 11. Nothing currently pending

No open user requests at the time this handoff was written. The dashboard, master Excel, and full ETL pipeline are all in sync and validated.
