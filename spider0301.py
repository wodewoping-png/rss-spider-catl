# spider_abstract_1216_wiley_recent5_keep_rss.py
# 目标：
# - daily 跑，只抓“昨天+前天”（UTC）做日报（所有源都一样，包括 Wiley）
# - Wiley 日期校验：对进入池的 Wiley 条目（已是昨天/前天RSS）再查 Crossref；没日期再查 PubMed
#   - 若校验日期存在且“在最近5天内”：保持 RSS 日期不动（不改时间、不drop）
#   - 若校验日期存在但“超过最近5天”：drop（认为 RSS 日期明显标错/条目过旧）
#   - 若无法校验日期：保持 RSS 日期不动（不改、不drop）
# - 导出 CSV 只保留 9 列：
#   title, link, source, published_str, pub_date, doi, abstract, abstract_source, must_have_abstract

import os
import re
import time
import random
from pathlib import Path
from datetime import datetime, timedelta, timezone
from html import unescape
from urllib.parse import urlsplit, urlunsplit
from xml.etree import ElementTree as ET

import feedparser
import pandas as pd
import requests
from playwright.sync_api import sync_playwright

# ================== 配置 ==================

FEED_LIST_FILE = Path("feeds1211.txt")
OUTPUT_DIR = Path("output")
OUTPUT_DIR.mkdir(exist_ok=True)

# 抓“昨天+前天”（UTC）
today_utc = datetime.now(timezone.utc).date()
yesterday_utc = today_utc - timedelta(days=1)
day_before_utc = today_utc - timedelta(days=2)
TARGET_DATES = {yesterday_utc, day_before_utc}

# ✅ Wiley 校验：canonical_date 在最近 N 天内，则保持 RSS 日期不动
WILEY_CANONICAL_RECENT_DAYS = 5

# 今日抓取基准时间（UTC 00:00）
TODAY_ANCHOR_DT = datetime(today_utc.year, today_utc.month, today_utc.day, tzinfo=timezone.utc)

# ✅ ScienceDirect 特殊源：pub_date 写“抓取日期的前一天”（UTC 00:00）
SD_ANCHOR_DT = TODAY_ANCHOR_DT - timedelta(days=1)
SD_ANCHOR_DATE = SD_ANCHOR_DT.date()

# 昨天CSV（用于复用&重试）
YESTERDAY_CSV = OUTPUT_DIR / f"news_with_abstract_{yesterday_utc.strftime('%Y-%m-%d')}.csv"
TODAY_CSV = OUTPUT_DIR / f"news_with_abstract_{today_utc.strftime('%Y-%m-%d')}.csv"

# API 配置
REQ_TIMEOUT = 15
API_SLEEP = 0.5
NCBI_TOOL = "literature_bot"
NCBI_EMAIL = os.getenv("NCBI_EMAIL", "qiaochuzhang@outlook.com")

# Playwright：CI 默认 headless
HEADLESS = os.getenv("PLAYWRIGHT_HEADLESS", "1") == "1"
IS_CI = (os.getenv("CI", "").lower() == "true")
BROWSER_CHANNEL = os.getenv("PLAYWRIGHT_CHANNEL", "").strip().lower()

# ✅ ScienceDirect 特殊源：发布日期不准 -> 直接抓最新 N 篇 + pub_date 写“抓取日期的前一天”
SD_FEED_APPLIED_ENERGY = "https://rss.sciencedirect.com/publication/science/03062619"
SD_FEED_ENERGY_POLICY = "https://rss.sciencedirect.com/publication/science/03014215"
SD_SPECIAL_LIMITS = {
    SD_FEED_APPLIED_ENERGY: 30,  # Applied Energy
    SD_FEED_ENERGY_POLICY: 7,    # Energy Policy
}
SD_SPECIAL_URLS = set(SD_SPECIAL_LIMITS.keys())

# ================== 正则与工具 ==================

TAG_RE = re.compile(r"<[^>]+>")
ALT_IMG_RE = re.compile(
    r'<img\b[^>]*\balt\s*=\s*(?P<q>["\'])(?P<alt>.*?)(?P=q)[^>]*>',
    flags=re.IGNORECASE | re.DOTALL
)
AVAILABLE_ONLINE_RE = re.compile(
    r"Available\s+online\s+(\d{1,2}\s+[A-Za-z]+\s+\d{4})",
    flags=re.IGNORECASE
)

MONTHS = {
    "january": 1, "february": 2, "march": 3, "april": 4, "may": 5, "june": 6,
    "july": 7, "august": 8, "september": 9, "october": 10, "november": 11, "december": 12
}

GENERIC_TITLES = {"graphical abstract", "table of contents", "toc", "cover image", "no title"}

ACS_TRASH_TITLES = {
    "issue editorial masthead",
    "issue publication information",
}

TITLE_EXCLUDE_KEYWORDS = ["editorial", "masthead", "issue information", "cover"]

# OUP(NSR) RSS description 抽摘要
OUP_ABS_RE = re.compile(
    r'boxTitle"\s*>\s*Abstract\s*<\s*/\s*div\s*>\s*(.*?)\s*(?:</span>|</description>|$)',
    flags=re.IGNORECASE | re.DOTALL
)

def clean_html_text(s: str) -> str:
    if not s:
        return ""
    s = unescape(s)
    s = TAG_RE.sub("", s)
    s = re.sub(r"\s+", " ", s)
    return s.strip()

def fix_mojibake(s: str) -> str:
    if not s:
        return s
    repl = {
        "鈥�": "'",
        "鈥": "'",
        "â": "'",
        "â": "-",
        "â": "-",
        "â": '"',
        "â": '"',
    }
    for k, v in repl.items():
        s = s.replace(k, v)
    return s

def looks_generic(title: str) -> bool:
    t = (title or "").strip().lower()
    if not t:
        return True
    return t in GENERIC_TITLES or len(t) < 5

def should_drop_by_title(title: str) -> bool:
    t = (title or "").strip().lower()
    if not t:
        return True
    if t in ACS_TRASH_TITLES:
        return True
    for kw in TITLE_EXCLUDE_KEYWORDS:
        if kw in t:
            return True
    return False

def extract_alt_from_html(html: str) -> str | None:
    if not html:
        return None
    m = ALT_IMG_RE.search(html)
    if not m:
        return None
    alt = clean_html_text(m.group("alt") or "")
    alt = fix_mojibake(alt)
    return alt or None

def parse_last_author_from_string(s: str) -> str:
    if not s:
        return ""
    s = fix_mojibake(clean_html_text(str(s)))
    if not s:
        return ""
    s = re.sub(r"\s+and\s+", ", ", s, flags=re.IGNORECASE)
    s = s.replace("&", ",")
    s = s.replace(" et al.", "").replace(" et al", "")
    parts = [p.strip() for p in s.split(",") if p.strip()]
    if not parts:
        return ""
    return parts[-1]

def extract_last_author_from_entry(entry) -> str:
    authors = entry.get("authors")
    if isinstance(authors, list) and authors:
        names = []
        for a in authors:
            if isinstance(a, dict):
                name = (a.get("name") or "").strip()
                if name:
                    names.append(name)
            else:
                name = str(a).strip()
                if name:
                    names.append(name)
        if names:
            return fix_mojibake(names[-1])

    author_detail = entry.get("author_detail")
    if isinstance(author_detail, dict):
        name = (author_detail.get("name") or "").strip()
        if name:
            return fix_mojibake(name)

    for key in ("dc_creator", "creator", "author"):
        val = entry.get(key)
        if isinstance(val, list) and val:
            if len(val) > 1:
                return fix_mojibake(str(val[-1]).strip())
            return parse_last_author_from_string(str(val[0]))
        if isinstance(val, str) and val.strip():
            return parse_last_author_from_string(val)

    return ""

def get_entry_title(entry) -> str:
    title_from_feed = fix_mojibake(clean_html_text(entry.get("title", "")))

    html_blobs = []
    if entry.get("content"):
        for c in entry.content:
            v = c.get("value") or ""
            if v:
                html_blobs.append(v)
    if entry.get("summary"):
        html_blobs.append(entry.get("summary"))

    alt_candidates = []
    for blob in html_blobs:
        alt_txt = extract_alt_from_html(blob)
        if alt_txt:
            alt_candidates.append(alt_txt)

    if looks_generic(title_from_feed) and alt_candidates:
        return max(alt_candidates, key=len)

    if title_from_feed and alt_candidates:
        best_alt = max(alt_candidates, key=len)
        if len(best_alt) >= len(title_from_feed) + 10 and not looks_generic(best_alt):
            return best_alt

    return title_from_feed or (alt_candidates[0] if alt_candidates else "")

def parse_date_strict(d: str):
    if not d:
        return None
    try:
        ts = pd.to_datetime(d, utc=True, errors="raise")
        return ts.to_pydatetime()
    except Exception:
        return None

def parse_available_online_date(description_html: str):
    if not description_html:
        return None
    desc_txt = clean_html_text(description_html)
    m = AVAILABLE_ONLINE_RE.search(desc_txt)
    if not m:
        return None
    date_str = m.group(1)
    parts = date_str.split()
    if len(parts) != 3:
        return parse_date_strict(date_str)
    day_s, month_s, year_s = parts
    try:
        day = int(day_s)
        month = MONTHS.get(month_s.lower())
        year = int(year_s)
        if not month:
            return parse_date_strict(date_str)
        return datetime(year, month, day, tzinfo=timezone.utc)
    except Exception:
        return parse_date_strict(date_str)

def get_entry_pub_date(entry):
    t = entry.get("published_parsed") or entry.get("updated_parsed")
    if t:
        return datetime(*t[:6], tzinfo=timezone.utc)

    dt = None
    for key in ("published", "updated"):
        dt = parse_date_strict(entry.get(key))
        if dt:
            break

    if not dt:
        desc = entry.get("summary") or ""
        dt = parse_available_online_date(desc)

    return dt

def extract_doi_from_url(url: str) -> str:
    if not url:
        return ""
    m = re.search(r"/10\.\d{4,9}/[^\s?#]+", url)
    if m:
        return m.group(0).lstrip("/")
    m2 = re.search(r"10\.\d{4,9}/[^\s?#]+", url)
    if m2:
        return m2.group(0)
    return ""

def extract_doi_from_entry(entry) -> str:
    for key in ("doi", "dc_identifier", "prism_doi", "dc:identifier", "id"):
        val = entry.get(key)
        if isinstance(val, list):
            for v in val:
                d = extract_doi_from_url(str(v))
                if d:
                    return d
        elif isinstance(val, str):
            d = extract_doi_from_url(val)
            if d:
                return d
    return extract_doi_from_url(entry.get("link") or "")

def normalize_link(url: str) -> str:
    if not url:
        return ""
    parts = urlsplit(url)
    return urlunsplit((parts.scheme, parts.netloc, parts.path, "", ""))

def record_key(doi: str, link: str) -> str:
    d = (doi or "").strip().lower()
    if d:
        return f"doi:{d}"
    return f"url:{normalize_link(link).lower()}"

def safe_get(url: str, params=None, headers=None):
    try:
        resp = requests.get(url, params=params, headers=headers, timeout=REQ_TIMEOUT)
        if resp.status_code == 200:
            return resp
        print(f"⚠️ 请求 {url} 状态码 {resp.status_code}")
        return None
    except Exception as e:
        print(f"⚠️ 请求 {url} 失败: {e}")
        return None

def within_days(pub_dt: datetime, days: int) -> bool:
    if not pub_dt:
        return False
    if pub_dt.tzinfo is None:
        pub_dt = pub_dt.replace(tzinfo=timezone.utc)
    return (today_utc - pub_dt.date()).days <= days

def in_target_dates(pub_dt: datetime | None) -> bool:
    if not pub_dt:
        return False
    if pub_dt.tzinfo is None:
        pub_dt = pub_dt.replace(tzinfo=timezone.utc)
    return pub_dt.date() in TARGET_DATES

# ================== Wiley 识别（用于日期校验） ==================

def is_wiley_record(source_title: str, link: str, doi: str) -> bool:
    ll = (link or "").lower()
    src = (source_title or "").lower()
    d = (doi or "").lower()
    return ("onlinelibrary.wiley.com" in ll) or ("wiley" in src) or d.startswith("10.1002/")

# ================== Publisher识别 + RSS抽摘要 ==================

def extract_oup_abstract_from_rss(desc_html: str) -> str:
    if not desc_html:
        return ""
    m = OUP_ABS_RE.search(desc_html)
    if not m:
        return ""
    return clean_html_text(m.group(1))

CELL_INPRESS_JOURNALS = {"chem", "joule", "oneear", "matter"}

def is_cellpress_inpress_any(source_title: str, link: str) -> bool:
    ll = (link or "").lower()
    if "cell.com" not in ll:
        return False
    m = re.search(r"cell\.com/([^/]+)/", ll)
    if not m:
        return False
    j = m.group(1).strip().lower()
    return j in CELL_INPRESS_JOURNALS

def is_oup_nsr(source_title: str, link: str) -> bool:
    src = (source_title or "").lower()
    ll = (link or "").lower()
    return ("national science review" in src) or ("academic.oup.com" in ll and "/nsr/" in ll)

def is_pnas(source_title: str, link: str) -> bool:
    src = (source_title or "").lower()
    ll = (link or "").lower()
    return ("pnas" in src) or ("pnas.org" in ll)

def is_acs_energy_letters(source_title: str, link: str) -> bool:
    src = (source_title or "").lower()
    ll = (link or "").lower()
    return ("acs energy letters" in src) or ("acsenergylett" in ll)

# ================== API 摘要 & 日期：Crossref / S2 / OpenAlex / PubMed ==================

def _crossref_pick_date(msg: dict) -> datetime | None:
    def parse_parts(obj):
        if not isinstance(obj, dict):
            return None
        parts = obj.get("date-parts")
        if not parts or not isinstance(parts, list) or not parts[0]:
            return None
        ymd = parts[0]
        try:
            y = int(ymd[0])
            m = int(ymd[1]) if len(ymd) >= 2 else 1
            d = int(ymd[2]) if len(ymd) >= 3 else 1
            return datetime(y, m, d, tzinfo=timezone.utc)
        except Exception:
            return None

    for k in ("published-online", "published-print", "issued", "created"):
        dt = parse_parts(msg.get(k))
        if dt:
            return dt
    return None

def _crossref_pick_last_author(msg: dict) -> str:
    authors = msg.get("author")
    if not isinstance(authors, list) or not authors:
        return ""
    last = authors[-1]
    if isinstance(last, dict):
        given = (last.get("given") or "").strip()
        family = (last.get("family") or "").strip()
        name = (last.get("name") or "").strip()
        if given or family:
            return (f"{given} {family}".strip())
        if name:
            return name
    return ""

def query_crossref_info(doi: str) -> dict:
    """Crossref：返回 abstract + canonical_pub_date（若有）"""
    url = f"https://api.crossref.org/works/{doi}"
    headers = {"User-Agent": f"ccus-bot (mailto:{NCBI_EMAIL})"}
    resp = safe_get(url, headers=headers)
    time.sleep(API_SLEEP)
    if not resp:
        return {}
    try:
        msg = resp.json().get("message", {})
    except Exception:
        return {}

    abstract = (msg.get("abstract", "") or "").strip()
    if abstract:
        abstract = re.sub(r"<[^>]+>", "", abstract).strip()

    pub_dt = _crossref_pick_date(msg)
    last_author = _crossref_pick_last_author(msg)
    return {
        "abstract": abstract,
        "source": "crossref",
        "canonical_pub_date": pub_dt,
        "last_author": last_author,
    }

def query_semanticscholar_abstract(doi: str) -> dict:
    url = f"https://api.semanticscholar.org/graph/v1/paper/DOI:{doi}"
    params = {"fields": "title,abstract,year,journal"}
    headers = {"User-Agent": f"ccus-bot (mailto:{NCBI_EMAIL})"}
    resp = safe_get(url, params=params, headers=headers)
    time.sleep(API_SLEEP)
    if not resp:
        return {}
    try:
        data = resp.json()
    except Exception:
        return {}
    return {"abstract": (data.get("abstract") or "").strip(), "source": "semanticscholar"}

def _openalex_reconstruct_abstract(inv_idx: dict) -> str:
    if not isinstance(inv_idx, dict) or not inv_idx:
        return ""
    pairs = []
    for w, poses in inv_idx.items():
        if not isinstance(poses, list):
            continue
        for p in poses:
            pairs.append((p, w))
    if not pairs:
        return ""
    pairs.sort(key=lambda x: x[0])
    return " ".join([w for _, w in pairs]).strip()

def query_openalex_abstract(doi: str) -> dict:
    url = f"https://api.openalex.org/works/https://doi.org/{doi}"
    headers = {"User-Agent": f"ccus-bot (mailto:{NCBI_EMAIL})"}
    resp = safe_get(url, headers=headers)
    time.sleep(API_SLEEP)
    if not resp:
        return {}
    try:
        data = resp.json()
    except Exception:
        return {}
    abstract = _openalex_reconstruct_abstract(data.get("abstract_inverted_index") or {})
    return {"abstract": (abstract or "").strip(), "source": "openalex"}

def query_openalex_last_author(doi: str) -> dict:
    url = f"https://api.openalex.org/works/https://doi.org/{doi}"
    headers = {"User-Agent": f"ccus-bot (mailto:{NCBI_EMAIL})"}
    resp = safe_get(url, headers=headers)
    time.sleep(API_SLEEP)
    if not resp:
        return {}
    try:
        data = resp.json()
    except Exception:
        return {}
    authorships = data.get("authorships") or []
    if not authorships:
        return {}
    last = authorships[-1]
    name = ""
    if isinstance(last, dict):
        author = last.get("author") or {}
        name = (author.get("display_name") or "").strip()
    return {"last_author": name, "source": "openalex"}

def query_semanticscholar_last_author(doi: str) -> dict:
    url = f"https://api.semanticscholar.org/graph/v1/paper/DOI:{doi}"
    params = {"fields": "authors"}
    headers = {"User-Agent": f"ccus-bot (mailto:{NCBI_EMAIL})"}
    resp = safe_get(url, params=params, headers=headers)
    time.sleep(API_SLEEP)
    if not resp:
        return {}
    try:
        data = resp.json()
    except Exception:
        return {}
    authors = data.get("authors") or []
    if not authors:
        return {}
    last = authors[-1]
    name = ""
    if isinstance(last, dict):
        name = (last.get("name") or "").strip()
    return {"last_author": name, "source": "semanticscholar"}

def get_last_author_via_apis(doi: str, crossref_cache: dict, openalex_cache: dict, s2_cache: dict) -> tuple[str, str]:
    if not doi:
        return "", ""

    info = crossref_cache.get(doi)
    if info is None:
        info = query_crossref_info(doi)
        crossref_cache[doi] = info
    last_author = (info.get("last_author") or "").strip()
    if last_author:
        return last_author, "crossref"

    info = openalex_cache.get(doi)
    if info is None:
        info = query_openalex_last_author(doi)
        openalex_cache[doi] = info
    last_author = (info.get("last_author") or "").strip()
    if last_author:
        return last_author, "openalex"

    info = s2_cache.get(doi)
    if info is None:
        info = query_semanticscholar_last_author(doi)
        s2_cache[doi] = info
    last_author = (info.get("last_author") or "").strip()
    if last_author:
        return last_author, "semanticscholar"

    return "", ""

def query_pubmed_info(doi: str) -> dict:
    """PubMed：返回 abstract + canonical_pub_date（若有）"""
    esearch_url = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi"
    esearch_params = {
        "db": "pubmed",
        "term": f"{doi}[DOI]",
        "retmode": "json",
        "tool": NCBI_TOOL,
        "email": NCBI_EMAIL,
    }
    resp = safe_get(esearch_url, params=esearch_params)
    time.sleep(API_SLEEP)
    if not resp:
        return {}
    try:
        data = resp.json()
        idlist = data.get("esearchresult", {}).get("idlist", [])
        if not idlist:
            return {}
        pmid = idlist[0]
    except Exception:
        return {}

    efetch_url = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi"
    efetch_params = {
        "db": "pubmed",
        "id": pmid,
        "retmode": "xml",
        "tool": NCBI_TOOL,
        "email": NCBI_EMAIL,
    }
    resp2 = safe_get(efetch_url, params=efetch_params)
    time.sleep(API_SLEEP)
    if not resp2:
        return {}
    try:
        root = ET.fromstring(resp2.text)
    except Exception:
        return {}

    article = root.find(".//PubmedArticle/MedlineCitation/Article")
    if article is None:
        return {}

    abstract = ""
    abs_list = article.findall("Abstract/AbstractText")
    if abs_list:
        parts = []
        for el in abs_list:
            text = "".join(el.itertext()).strip()
            if text:
                parts.append(text)
        abstract = "\n".join(parts).strip()

    pub_dt = None
    # 优先电子发表日期
    for ad in root.findall(".//PubmedArticle/MedlineCitation/Article/ArticleDate"):
        dtype = (ad.get("DateType") or "").lower()
        if dtype == "electronic":
            y = ad.findtext("Year") or ""
            m = ad.findtext("Month") or "1"
            d = ad.findtext("Day") or "1"
            try:
                pub_dt = datetime(int(y), int(m), int(d), tzinfo=timezone.utc)
                break
            except Exception:
                pass

    # 回退到期刊 PubDate
    if not pub_dt:
        pd_el = root.find(".//PubmedArticle/MedlineCitation/Article/Journal/JournalIssue/PubDate")
        if pd_el is not None:
            y = pd_el.findtext("Year")
            m = pd_el.findtext("Month") or "1"
            d = pd_el.findtext("Day") or "1"
            if y:
                try:
                    mm = m
                    if isinstance(mm, str) and mm.isalpha():
                        mm = MONTHS.get(mm.lower(), 1)
                    pub_dt = datetime(int(y), int(mm), int(d), tzinfo=timezone.utc)
                except Exception:
                    pass

    return {"abstract": abstract, "source": "pubmed", "canonical_pub_date": pub_dt}

def get_abstract_via_apis(doi: str, aggressive: bool = False) -> tuple[str, str]:
    if not doi:
        return "", ""
    print(f"   [API] DOI={doi} (aggressive={aggressive})")

    if aggressive:
        order = (query_semanticscholar_abstract, query_openalex_abstract, query_crossref_info, query_pubmed_info)
    else:
        order = (query_crossref_info, query_semanticscholar_abstract, query_openalex_abstract, query_pubmed_info)

    for fn in order:
        info = fn(doi)
        abs_txt = (info.get("abstract") or "").strip()
        if abs_txt:
            return abs_txt, info.get("source", "")
    return "", ""

# ================== Playwright：仅 Nature / RSC ==================

def maybe_human_like_wait():
    if IS_CI:
        return
    time.sleep(random.uniform(0.6, 1.6))

def extract_nature_abstract(page) -> str:
    try:
        page.wait_for_timeout(500)
        el = page.query_selector('section[data-title="Abstract"] .c-article-section__content')
        if not el:
            el = page.query_selector('section[aria-labelledby="Abs1"] .c-article-section__content')
        if not el:
            return ""
        return clean_html_text(el.inner_html())
    except Exception:
        return ""

def extract_rsc_abstract(page) -> str:
    try:
        page.wait_for_timeout(500)
        el = page.query_selector("div.capsule__text")
        if not el:
            el = page.query_selector("h3.article-abstract__heading + div.capsule__column-wrapper div.capsule__text")
        if not el:
            return ""
        return clean_html_text(el.inner_html())
    except Exception:
        return ""

def get_html_abstract_for_record(page, link: str) -> tuple[str, str]:
    ll = (link or "").lower()
    if not ll:
        return "", ""
    is_nature = "nature.com" in ll
    is_rsc = "rsc.org" in ll or "pubs.rsc.org" in ll
    if not (is_nature or is_rsc):
        return "", ""

    print(f"   [HTML] {link}")
    try:
        maybe_human_like_wait()
        page.goto(link, wait_until="load", timeout=60_000)
        maybe_human_like_wait()
    except Exception as e:
        print(f"   ⚠️ HTML打开失败: {e}")
        return "", ""

    if is_nature:
        abs_txt = extract_nature_abstract(page)
        return (abs_txt, "nature_html") if abs_txt else ("", "")
    if is_rsc:
        abs_txt = extract_rsc_abstract(page)
        return (abs_txt, "rsc_html") if abs_txt else ("", "")
    return "", ""

def launch_browser(p):
    kwargs = dict(headless=HEADLESS, slow_mo=0 if HEADLESS else random.randint(50, 120))
    if BROWSER_CHANNEL in ("chrome", "msedge", "edge"):
        kwargs["channel"] = "chrome" if BROWSER_CHANNEL == "chrome" else "msedge"
    return p.chromium.launch(**kwargs)

# ================== 昨日缓存 ==================

def load_yesterday_cache(path: Path):
    prev_by_key = {}
    prev_has_abs_keys = set()
    prev_no_abs_recent3_keys = set()

    if not path.exists():
        print(f"ℹ️ 昨日CSV不存在：{path}")
        return prev_by_key, prev_has_abs_keys, prev_no_abs_recent3_keys

    df = pd.read_csv(path, encoding="utf-8-sig", keep_default_na=False)
    if df.empty:
        return prev_by_key, prev_has_abs_keys, prev_no_abs_recent3_keys

    df["pub_date"] = pd.to_datetime(df.get("pub_date", ""), utc=True, errors="coerce")

    for _, row in df.iterrows():
        doi = str(row.get("doi", "") or "").strip()
        link = str(row.get("link", "") or "").strip()
        key = record_key(doi, link)

        pub_dt = row.get("pub_date")
        pub_dt = pub_dt.to_pydatetime() if hasattr(pub_dt, "to_pydatetime") else None

        abstract = str(row.get("abstract", "") or "").strip()

        rec = {
            "title": str(row.get("title", "") or ""),
            "link": link,
            "source": str(row.get("source", "") or ""),
            "published_str": str(row.get("published_str", "") or ""),
            "pub_date": pub_dt,
            "doi": doi,
            "last_author": str(row.get("last_author", "") or ""),
            "last_author_source": str(row.get("last_author_source", "") or ""),
            "author_needs_api": False,
            "abstract": abstract,
            "abstract_source": str(row.get("abstract_source", "") or ""),
            "must_have_abstract": False,
        }
        prev_by_key[key] = rec

        if abstract:
            prev_has_abs_keys.add(key)
        else:
            if pub_dt and within_days(pub_dt, 3):
                prev_no_abs_recent3_keys.add(key)

    print(f"✅ 昨日缓存：{len(prev_by_key)} 条；有摘要 {len(prev_has_abs_keys)}；需重试 {len(prev_no_abs_recent3_keys)}")
    return prev_by_key, prev_has_abs_keys, prev_no_abs_recent3_keys

# ================== RSS 收集 ==================

def read_feed_list(path: Path) -> list[str]:
    if not path.exists():
        raise FileNotFoundError(f"RSS 源文件不存在: {path}")
    with path.open("r", encoding="utf-8") as f:
        return [line.strip() for line in f if line.strip()]

def collect_rss_records(prev_by_key, prev_has_abs_keys, prev_no_abs_recent3_keys):
    urls = read_feed_list(FEED_LIST_FILE)
    today_records = {}

    for feed_url in urls:
        print(f"\n📡 RSS: {feed_url}")
        feed = feedparser.parse(feed_url)
        source_title = feed.feed.get("title", feed_url)

        is_sd_special = feed_url in SD_SPECIAL_URLS
        special_limit = SD_SPECIAL_LIMITS.get(feed_url, 0)
        special_taken = 0

        if is_sd_special:
            print(f"  ✅ ScienceDirect特殊策略：抓最新 {special_limit} 条；pub_date 统一写抓取日期的前一天 {SD_ANCHOR_DATE} (UTC)")

        for entry in feed.entries:
            # ========== A) ScienceDirect 特殊源：只抓最新 N 条，不做日期过滤 ==========
            if is_sd_special:
                if special_taken >= special_limit:
                    break
                pub_date = SD_ANCHOR_DT
                published_str = pub_date.strftime("%Y-%m-%d %H:%M:%S %Z")
            # ========== B) 普通源：严格按昨天+前天过滤（包括 Wiley） ==========
            else:
                pub_date = get_entry_pub_date(entry)
                if not pub_date or pub_date.date() not in TARGET_DATES:
                    continue
                published_str = pub_date.strftime("%Y-%m-%d %H:%M:%S %Z")

            title = get_entry_title(entry)
            if should_drop_by_title(title):
                continue

            link = entry.get("link") or ""
            desc_html = entry.get("summary") or entry.get("description") or ""
            doi = extract_doi_from_entry(entry)
            key = record_key(doi, link)

            # 昨天已有摘要 -> 直接复用
            if key in prev_has_abs_keys:
                today_records[key] = {**prev_by_key[key], "must_have_abstract": False}
                # 保持 pub_date / published_str 为当前 RSS 的（若昨日缺）
                if today_records[key].get("pub_date") is None and pub_date is not None:
                    today_records[key]["pub_date"] = pub_date
                if not (today_records[key].get("published_str") or "").strip() and published_str:
                    today_records[key]["published_str"] = published_str
                if is_sd_special:
                    special_taken += 1
                continue

            abstract = ""
            abstract_source = ""

            last_author = ""
            last_author_source = ""
            author_needs_api = False

            if is_oup_nsr(source_title, link) or is_pnas(source_title, link):
                author_needs_api = True
            else:
                last_author = extract_last_author_from_entry(entry)
                if last_author:
                    last_author_source = "rss"
                else:
                    author_needs_api = True

            # ✅ Cell：RSS description 就是摘要
            if is_cellpress_inpress_any(source_title, link):
                abstract = clean_html_text(desc_html)
                abstract_source = "rss_cell"

            # ✅ OUP(NSR)：RSS description 内含 Abstract
            elif is_oup_nsr(source_title, link):
                abs_txt = extract_oup_abstract_from_rss(desc_html)
                if abs_txt:
                    abstract = abs_txt
                    abstract_source = "rss_oup"

            must_have_abstract = key in prev_no_abs_recent3_keys

            today_records[key] = {
                "title": title,
                "link": link,
                "source": source_title,
                "published_str": published_str,
                "pub_date": pub_date,
                "doi": doi,
                "last_author": (last_author or "").strip(),
                "last_author_source": (last_author_source or "").strip(),
                "author_needs_api": author_needs_api,
                "abstract": (abstract or "").strip(),
                "abstract_source": (abstract_source or "").strip(),
                "must_have_abstract": must_have_abstract,
                # ✅ Wiley 校验需要：保存 RSS 原始日期（仅内存使用，不导出）
                "rss_pub_date": pub_date,
                "rss_published_str": published_str,
                "_drop": False,
            }

            if is_sd_special:
                special_taken += 1

        if is_sd_special:
            print(f"  ✅ 特殊源实际收录：{special_taken}/{special_limit} 条（标题过滤/去重后可能少于limit）")

    return today_records

def carry_over_prev_with_abstract(today_records, prev_by_key, prev_has_abs_keys):
    moved = 0
    for key in prev_has_abs_keys:
        rec = prev_by_key.get(key)
        if not rec:
            continue
        pub_dt = rec.get("pub_date")
        if pub_dt and within_days(pub_dt, 2):
            if key not in today_records:
                today_records[key] = {**rec, "must_have_abstract": False}
                moved += 1
    print(f"✅ 搬运昨日（2天内且有摘要）：{moved} 条")

def build_retry_records(prev_by_key, prev_no_abs_recent3_keys, today_records):
    retry = []
    for key in prev_no_abs_recent3_keys:
        if key in today_records:
            continue
        rec = prev_by_key[key].copy()
        rec["must_have_abstract"] = True
        if rec.get("pub_date"):
            retry.append(rec)
    print(f"✅ 重试列表（昨日3天内无摘要 & 今日RSS未出现）：{len(retry)} 条")
    return retry

# ================== 摘要补全 + Wiley 日期复核（recent5 规则） ==================

def enrich_with_html_then_api(records: list[dict]):
    need_html = any(
        (("nature.com" in (r.get("link","").lower())) or ("rsc.org" in (r.get("link","").lower())) or ("pubs.rsc.org" in (r.get("link","").lower())))
        and not (r.get("abstract") or "").strip()
        for r in records
    )

    if need_html:
        print("\n🔧 HTML阶段：抓 Nature/RSC（headless适配）...")
        with sync_playwright() as p:
            browser = launch_browser(p)
            context = browser.new_context(
                user_agent=(
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/123.0.0.0 Safari/537.36"
                ),
                viewport={"width": 1400, "height": 900},
            )
            page = context.new_page()

            for r in records:
                if (r.get("abstract") or "").strip():
                    continue
                ll = (r.get("link") or "").lower()
                if not ("nature.com" in ll or "rsc.org" in ll or "pubs.rsc.org" in ll):
                    continue
                abs_txt, src = get_html_abstract_for_record(page, r.get("link", ""))
                if abs_txt:
                    r["abstract"] = abs_txt.strip()
                    r["abstract_source"] = src

            browser.close()

    crossref_cache: dict[str, dict] = {}
    pubmed_cache: dict[str, dict] = {}
    openalex_cache: dict[str, dict] = {}
    s2_cache: dict[str, dict] = {}

    print(f"\n🔧 API阶段：补全剩余摘要 + Wiley日期复核（canonical在最近{WILEY_CANONICAL_RECENT_DAYS}天内则不改RSS；否则drop）...")
    for r in records:
        doi = (r.get("doi") or "").strip()

        # Wiley 日期复核（仅对已入池的 Wiley：即RSS已是昨天/前天）
        if doi and is_wiley_record(r.get("source",""), r.get("link",""), doi):
            # 保证 RSS 备份字段存在
            if r.get("rss_pub_date") in ("", None):
                r["rss_pub_date"] = r.get("pub_date")
            if not (r.get("rss_published_str") or "").strip():
                r["rss_published_str"] = r.get("published_str","") or ""

            rss_dt = r.get("rss_pub_date")

            # 1) Crossref
            info = crossref_cache.get(doi)
            if info is None:
                info = query_crossref_info(doi)
                crossref_cache[doi] = info
            canonical_dt = info.get("canonical_pub_date")

            # 2) PubMed fallback
            if not canonical_dt:
                info2 = pubmed_cache.get(doi)
                if info2 is None:
                    info2 = query_pubmed_info(doi)
                    pubmed_cache[doi] = info2
                canonical_dt = info2.get("canonical_pub_date")

            # 3) recent5 规则：不改RSS；若明显过旧则 drop
            if canonical_dt:
                if within_days(canonical_dt, WILEY_CANONICAL_RECENT_DAYS):
                    # ✅ canonical最近5天：保持 RSS 日期不动
                    r["pub_date"] = rss_dt
                    r["published_str"] = r.get("rss_published_str") or ""
                    r["_drop"] = False
                else:
                    # ✅ canonical超过5天：drop
                    r["_drop"] = True
            else:
                # ✅ 无法校验：保持 RSS 日期不动
                r["pub_date"] = rss_dt
                r["published_str"] = r.get("rss_published_str") or ""
                r["_drop"] = False

            # 可选：顺手用 Crossref abstract（如果你还没摘要）
            if not (r.get("abstract") or "").strip():
                abs_txt = (info.get("abstract") or "").strip()
                if abs_txt:
                    r["abstract"] = abs_txt
                    r["abstract_source"] = "crossref"

            if not (r.get("last_author") or "").strip() and (r.get("author_needs_api") or is_oup_nsr(r.get("source",""), r.get("link","")) or is_pnas(r.get("source",""), r.get("link",""))):
                last_author = (info.get("last_author") or "").strip()
                if last_author:
                    r["last_author"] = last_author
                    r["last_author_source"] = "crossref"

        # 作者补全：不依赖摘要是否已存在
        if doi and not (r.get("last_author") or "").strip() and (r.get("author_needs_api") or is_oup_nsr(r.get("source",""), r.get("link","")) or is_pnas(r.get("source",""), r.get("link",""))):
            last_author, a_src = get_last_author_via_apis(doi, crossref_cache, openalex_cache, s2_cache)
            if last_author:
                r["last_author"] = last_author
                r["last_author_source"] = a_src

        # 原有：补全剩余摘要（先 HTML 后 API）
        if (r.get("abstract") or "").strip():
            continue

        if not doi:
            continue

        aggressive = is_acs_energy_letters(r.get("source",""), r.get("link",""))
        abs_txt, src = get_abstract_via_apis(doi, aggressive=aggressive)
        if abs_txt:
            r["abstract"] = abs_txt.strip()
            r["abstract_source"] = src

# ================== 导出（只保留9列） ==================

def export_records(today_records: dict):
    if not today_records:
        print("⚠️ 没有记录可导出。")
        return

    def _record_fill_score(r: dict) -> int:
        fields = [
            "title",
            "link",
            "source",
            "published_str",
            "pub_date",
            "doi",
            "last_author",
            "abstract",
            "abstract_source",
        ]
        score = 0
        for f in fields:
            v = r.get(f)
            if v is None:
                continue
            if isinstance(v, str):
                if v.strip():
                    score += 1
            else:
                score += 1
        return score

    def _merge_records_by_link(records: list[dict]) -> list[dict]:
        groups: dict[str, list[dict]] = {}
        for r in records:
            link = (r.get("link") or "").strip()
            norm = normalize_link(link).lower() if link else ""
            key = f"url:{norm}" if norm else record_key(r.get("doi", ""), link)
            groups.setdefault(key, []).append(r)

        merged = []
        for _, items in groups.items():
            if len(items) == 1:
                merged.append(items[0])
                continue
            items = sorted(items, key=_record_fill_score, reverse=True)
            base = items[0].copy()
            for other in items[1:]:
                for f in ("title", "link", "source", "published_str", "doi", "last_author", "abstract", "abstract_source"):
                    if not (base.get(f) or "").strip():
                        v = (other.get(f) or "").strip() if isinstance(other.get(f), str) else other.get(f)
                        if isinstance(v, str):
                            if v.strip():
                                base[f] = v
                        elif v is not None:
                            base[f] = v
                if base.get("pub_date") is None and other.get("pub_date") is not None:
                    base["pub_date"] = other.get("pub_date")
                base["must_have_abstract"] = bool(base.get("must_have_abstract")) or bool(other.get("must_have_abstract"))
            merged.append(base)
        return merged

    merged_records = _merge_records_by_link(list(today_records.values()))
    df_all = pd.DataFrame(merged_records)

    keep_cols = [
        "title",
        "link",
        "source",
        "published_str",
        "pub_date",
        "doi",
        "last_author",
        "abstract",
        "abstract_source",
        "must_have_abstract",
    ]

    # 只选白名单列，避免多余列名出现在CSV
    df = df_all.loc[:, [c for c in keep_cols if c in df_all.columns]]
    for c in keep_cols:
        if c not in df.columns:
            df[c] = ""

    df["pub_date"] = pd.to_datetime(df["pub_date"], utc=True, errors="coerce")

    for col in ["title", "link", "source", "published_str", "doi", "last_author", "abstract", "abstract_source"]:
        df[col] = df[col].fillna("").astype(str)

    df["must_have_abstract"] = df["must_have_abstract"].fillna(False).astype(bool).astype(int)

    df = df.sort_values(by="pub_date")
    df = df[keep_cols]

    df.to_csv(TODAY_CSV, index=False, encoding="utf-8-sig")
    print(f"\n✅ 导出：{len(df)} 条 -> {TODAY_CSV}")

# ================== 主流程 ==================

def main():
    print("▶ spider_abstract_1216_wiley_recent5_keep_rss.py 启动")
    print(f"▶ CI={IS_CI}, HEADLESS={HEADLESS}, CHANNEL='{BROWSER_CHANNEL or 'playwright-chromium'}'")
    print(f"▶ 目标日期(UTC)：{sorted(TARGET_DATES)}")
    print(f"▶ Wiley复核规则：canonical_date 在最近{WILEY_CANONICAL_RECENT_DAYS}天内 => 保持RSS日期；否则 drop")
    print(f"▶ 昨日CSV：{YESTERDAY_CSV}")
    print(f"▶ 今日输出：{TODAY_CSV}")
    print("▶ ScienceDirect 特殊源限制 + pub_date写抓取日期的前一天：")
    for k, v in SD_SPECIAL_LIMITS.items():
        print(f"   - {k} -> latest {v}, pub_date={SD_ANCHOR_DATE} (UTC)")

    prev_by_key, prev_has_abs_keys, prev_no_abs_recent3_keys = load_yesterday_cache(YESTERDAY_CSV)

    today_records = collect_rss_records(prev_by_key, prev_has_abs_keys, prev_no_abs_recent3_keys)

    carry_over_prev_with_abstract(today_records, prev_by_key, prev_has_abs_keys)

    records_list = list(today_records.values())
    enrich_with_html_then_api(records_list)

    # 重建 dict（以 key 去重）
    today_records = {record_key(r.get("doi",""), r.get("link","")): r for r in records_list}

    # ✅ Wiley：只根据 _drop 标记丢弃（不会扩大抓取范围）
    drop_keys = [k for k, r in today_records.items() if r.get("_drop")]
    for k in drop_keys:
        today_records.pop(k, None)
    if drop_keys:
        print(f"🧹 Wiley复核：canonical_date 非最近{WILEY_CANONICAL_RECENT_DAYS}天 -> 丢弃 {len(drop_keys)} 条")

    # 重试：昨日3天内无摘要、今日RSS未出现
    retry_records = build_retry_records(prev_by_key, prev_no_abs_recent3_keys, today_records)
    if retry_records:
        enrich_with_html_then_api(retry_records)
        added = 0
        for r in retry_records:
            # 必须有摘要才加入
            if not (r.get("abstract") or "").strip():
                continue
            # Wiley 若被标记 drop，不加入
            if r.get("_drop"):
                continue
            k = record_key(r.get("doi",""), r.get("link",""))
            today_records[k] = r
            added += 1
        print(f"✅ 重试成功加入今日：{added} 条（失败的不加入 / Wiley过旧的不加入）")

    # must_have_abstract 且无摘要 -> 丢弃
    drop = [k for k, r in today_records.items() if r.get("must_have_abstract") and not (r.get("abstract") or "").strip()]
    for k in drop:
        today_records.pop(k, None)
    if drop:
        print(f"🧹 丢弃 must_have_abstract 且无摘要：{len(drop)} 条")

    export_records(today_records)

if __name__ == "__main__":
    main()
