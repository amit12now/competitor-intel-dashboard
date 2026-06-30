# -*- coding: utf-8 -*-
"""
Shared taxonomy for the Competitive Intelligence ETL pipeline:
- tracked competitor brands + alias matching
- "our brand" (Atlas Copco) alias matching
- theme keyword dictionary
- sentiment lexicon
- geography keyword list
- theme -> (possible meaning template, team to review) mapping
"""

OUR_BRAND = "Atlas Copco"
OUR_BRAND_ALIASES = ["atlas copco", "atlascopco"]

COMPETITORS = {
    "Ingersoll Rand": [
        "ingersoll rand", "ingersoll-rand", "ingersoll rand compressor systems",
        "ingersollrand", "ingersoll rand air compressors", "vicente reynal",
    ],
    "Sullair": ["sullair"],
    "ELGi": [
        "elgi", "elgi equipments", "elgi north america", "elgi air compressors",
        "jairam varadaraj",
    ],
    "FS-Curtis": ["fs-curtis", "fs curtis", "fscurtis", "toledo tools"],
    "Quincy Compressor": ["quincy compressor", "quincy "],
    "Kaeser Compressors": ["kaeser", "kaeserusa", "kaeser usa", "kaeser compressors"],
    "BOGE": [
        "boge kompressoren", "bogekompressoren", " boge ", "boge en compressed air",
        "boge e-series",
    ],
    "Hitachi Global Air Power": ["hitachi global air power", "hideki fujimoto"],
    "Gardner Denver": ["gardner denver"],
}

COMPETITOR_LIST = list(COMPETITORS.keys())

ALIAS_TO_CANON = {}
for canon, aliases in COMPETITORS.items():
    for a in aliases:
        ALIAS_TO_CANON[a.strip().lower()] = canon

GENERIC_INDUSTRY_KEYWORDS = [
    "air compressor", "compressed air", "rotary screw", "oil-free compressor",
    "oil free compressor", "reciprocating compressor", "centrifugal compressor",
    "compressor systems", "industrial air", "psi", "cfm air", "pneumatic",
]


def find_companies_in_text(text):
    if not text:
        return set()
    t = " " + text.lower() + " "
    found = set()
    for alias, canon in ALIAS_TO_CANON.items():
        if alias in t:
            found.add(canon)
    return found


def mentions_our_brand(text):
    if not text:
        return False
    t = text.lower()
    return any(a in t for a in OUR_BRAND_ALIASES)


def mentions_industry_keyword(text):
    if not text:
        return False
    t = text.lower()
    return any(k in t for k in GENERIC_INDUSTRY_KEYWORDS)


THEMES = [
    "Energy efficiency", "Service / aftermarket", "Maintenance / repair",
    "Oil-free compressors", "Rotary screw compressors", "Rental",
    "Sustainability", "Industrial productivity", "New product launch",
    "Customer success / case study", "Training / webinar",
    "Dealer / distributor", "Hiring / expansion", "Local market activity",
]

THEME_KEYWORDS = {
    "Energy efficiency": [
        "energy efficien", "energy sav", "kwh", "lower energy", "power consumption",
        "energy cost", "vsd", "variable speed drive", "reduce energy",
    ],
    "Service / aftermarket": [
        "aftermarket", "service plan", "care services", "service contract",
        "genuine parts", "service technician", "service network", "spare parts",
        "service agreement", "premium aftermarket",
    ],
    "Maintenance / repair": [
        "maintenance", "repair", "overhaul", "rebuild", "downtime", "preventive",
        "tune-up", "tune up", "inspection",
    ],
    "Oil-free compressors": ["oil-free", "oil free", "oilless", "oil-less"],
    "Rotary screw compressors": ["rotary screw", "screw compressor", "screw air compressor"],
    "Rental": ["rental", "rent a compressor", "temporary air", "fleet rental", "rental fleet"],
    "Sustainability": [
        "sustainab", "esg", "carbon neutral", "carbon footprint", "net zero",
        "green energy", "emission", "environment", "world environment day",
        "renewable",
    ],
    "Industrial productivity": [
        "productivity", "uptime", "throughput", "operational efficiency",
        "total cost of ownership", "manufacturing efficiency", "lean on us",
    ],
    "New product launch": [
        "launch", "introduc", "new product", "unveil", "new series", "new model",
        "now available", "next generation", "next-gen",
    ],
    "Customer success / case study": [
        "case study", "customer success", "testimonial", "client story",
        "customer testimonial", "success story",
    ],
    "Training / webinar": [
        "webinar", "training", "workshop", "certification", "online course",
        "virtual expo", "masterclass",
    ],
    "Dealer / distributor": [
        "dealer", "distributor", "channel partner", "dealer network",
        "authorized dealer", "partner network",
    ],
    "Hiring / expansion": [
        "hiring", "we're hiring", "we are hiring", "job opening", "careers",
        "join our team", "expansion", "groundbreaking", "new facility",
        "campus expansion", "acquires", "acquisition", "new plant", "expand",
    ],
    "Local market activity": [
        "expo", "trade show", "exhibition", "networking meet", "conference",
        "summit", "regional", "local market",
    ],
}


def classify_themes(text, max_themes=2):
    if not text:
        return []
    t = text.lower()
    scores = []
    for theme, kws in THEME_KEYWORDS.items():
        score = sum(t.count(kw) for kw in kws)
        if score > 0:
            scores.append((score, theme))
    scores.sort(key=lambda x: -x[0])
    return [theme for _, theme in scores[:max_themes]]


POSITIVE_WORDS = [
    "great", "excellent", "proud", "excited", "award", "celebrate", "congrat",
    "thrilled", "amazing", "fantastic", "love", "trust", "reliable", "best",
    "innovative", "success", "happy", "honored", "grateful", "milestone",
    "outstanding", "impressive", "recommend", "winning", "win", "thank you",
]
NEGATIVE_WORDS = [
    "recall", "issue", "complaint", "fail", "failure", "problem", "broken",
    "disappoint", "delay", "poor", "bad", "concern", "lawsuit", "defect",
    "shortage", "downtime", "leak", "breakdown", "frustrat", "unhappy",
]


def classify_sentiment(text):
    if not text:
        return "Neutral"
    t = text.lower()
    pos = sum(t.count(w) for w in POSITIVE_WORDS)
    neg = sum(t.count(w) for w in NEGATIVE_WORDS)
    if pos > 0 and neg > 0:
        return "Mixed"
    if pos > neg and pos > 0:
        return "Positive"
    if neg > pos and neg > 0:
        return "Negative"
    return "Neutral"


GEO_KEYWORDS = [
    "United States", "USA", "U.S.", "North America", "Canada", "Mexico",
    "India", "Europe", "Germany", "United Kingdom", "UK", "France", "Italy",
    "Spain", "Brazil", "China", "Japan", "Southeast Asia", "Middle East",
    "Africa", "Australia", "Michigan City", "Charlotte", "St. Louis",
    "Davidson", "Fredericksburg", "Visakhapatnam", "Haridwar", "Coimbatore",
    "Texas", "California", "Ohio", "Illinois", "Wisconsin", "Georgia",
    "Pennsylvania", "New York", "Indiana", "Tennessee",
]


def find_geography(text):
    if not text:
        return ""
    found = []
    for g in GEO_KEYWORDS:
        if g.lower() in text.lower():
            found.append(g)
    seen = set()
    out = []
    for g in found:
        if g not in seen:
            out.append(g)
            seen.add(g)
    return "; ".join(out[:3])


THEME_INSIGHT = {
    "Energy efficiency": (
        "May indicate continued positioning around energy savings messaging; could be worth reviewing for content/SEO angle comparisons.",
        "SEO / Content",
    ),
    "Service / aftermarket": (
        "This may indicate a competitor focus on aftermarket/service demand; potential opportunity for sales enablement review.",
        "Sales / Product Marketing",
    ),
    "Maintenance / repair": (
        "Could indicate ongoing service/maintenance positioning; relevant for product marketing or content gap review.",
        "Product Marketing / Content",
    ),
    "Oil-free compressors": (
        "Potential opportunity to compare oil-free product messaging; may be useful for product marketing review.",
        "Product Marketing",
    ),
    "Rotary screw compressors": (
        "Core product-line activity; may be relevant for product marketing/sales battlecard review.",
        "Product Marketing / Sales",
    ),
    "Rental": (
        "Could indicate rental/fleet market activity; may be worth reviewing with sales for regional rental demand signals.",
        "Sales",
    ),
    "Sustainability": (
        "May reflect sustainability/ESG messaging push; potential opportunity for content or PR angle comparison.",
        "PR / Content",
    ),
    "Industrial productivity": (
        "General productivity/efficiency messaging; may be relevant for competitive content review.",
        "Content / SEO",
    ),
    "New product launch": (
        "Possible new product/service activity; this page or post may be relevant for SEO gap analysis and product marketing review.",
        "Product Marketing / SEO",
    ),
    "Customer success / case study": (
        "Customer proof point; could be useful for sales enablement or content benchmarking review.",
        "Sales / Content",
    ),
    "Training / webinar": (
        "Thought leadership / training activity; may be relevant for content or demand-gen review.",
        "Content / Marketing",
    ),
    "Dealer / distributor": (
        "May indicate channel/distributor network activity; potential relevance for sales/channel team discussion.",
        "Sales",
    ),
    "Hiring / expansion": (
        "Multiple roles or facility activity may suggest capacity or regional expansion; could be worth reviewing with regional sales/marketing.",
        "Leadership / Sales",
    ),
    "Local market activity": (
        "Event/market participation signal; may be relevant for regional marketing or PR awareness review.",
        "Marketing / PR",
    ),
    "General brand activity": (
        "General brand visibility activity; logged for completeness, likely lower priority for action.",
        "Marketing",
    ),
}
