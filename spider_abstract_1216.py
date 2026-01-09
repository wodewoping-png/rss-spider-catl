# spider_abstract_1216_wiley_datecheck_keep_rss_tolerance.py
# 逻辑：
# - daily 跑，目标日期仍是昨天+前天（UTC）
# - Wiley：先放行入池；后续对每条做 Crossref 日期校验；Crossref 无日期再查 PubMed
# - “容差规则”：若(校验日期 - RSS日期)的天数差 <= 3，则保持 RSS 日期为准（不改时间、不drop）
#             若差 > 3，则以校验日期为准决定去留：不在 TARGET_DATES -> drop；在 -> 用校验日期覆盖
# - 若无法校验（Crossref+PubMed都无日期），则用 RSS 日期决定去留（不改时间）

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

# 容差：校验日期与RSS日期相差 <= N 天，则仍以RSS为准
WILEY_DATE_TOLERANCE_DAYS = 3

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

def date_diff_days(a: datetime | None, b: datetime | None) -> int | None:
    """返回两个日期(date层面)的绝对天数差；任一为空则返回 None"""
    if not a or not b:
        return None
    da = a.date() if hasattr(a, "date") else None
    db = b.date() if hasattr(b, "date") else None
    if not da or not db:
        return None
    return abs((da - db).days)

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

def query_crossref_abstract(doi: str) -> dict:
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
    return {"abstract": abstract, "source": "crossref", "canonical_pub_date": pub_dt}

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

def query_pubmed_abstract(doi: str) -> dict:
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
        order = (query_semanticscholar_abstract, query_openalex_abstract, query_crossref_abstract, query_pubmed_abstract)
    else:
        order = (query_crossref_abstract, query_semanticscholar_abstract, query_openalex_abstract, query_pubmed_abstract)

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
            "abstract": abstract,
            "abstract_source": str(row.get("abstract_source", "") or ""),
            "must_have_abstract": False,
            "pub_date_source": str(row.get("pub_date_source", "") or ""),
            "rss_pub_date": "",
            "rss_published_str": "",
            "wiley_date_diff_days": "",
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
        elif feed_url in (SD_FEED_APPLIED_ENERGY, SD_FEED_ENERGY_POLICY):
            print("  ✅ ScienceDirect特殊策略启用")

        for entry in feed.entries:
            # ========== A) ScienceDirect 特殊源：只抓最新 N 条，不做日期过滤 ==========
            if is_sd_special:
                if special_taken >= special_limit:
                    break
                pub_date = SD_ANCHOR_DT
                published_str = pub_date.strftime("%Y-%m-%d %H:%M:%S %Z")
                link = entry.get("link") or ""
                doi = extract_doi_from_url(link)

            # ========== B) 普通源：按昨天+前天过滤（但 Wiley 放行；无日期也不占位） ==========
            else:
                pub_date = get_entry_pub_date(entry)
                link = entry.get("link") or ""
                doi = extract_doi_from_url(link)

                if not is_wiley_record(source_title, link, doi):
                    if not pub_date or pub_date.date() not in TARGET_DATES:
                        continue
                else:
                    # Wiley 放行
                    pass

                published_str = pub_date.strftime("%Y-%m-%d %H:%M:%S %Z") if pub_date else ""

            title = get_entry_title(entry)
            if should_drop_by_title(title):
                continue

            desc_html = entry.get("summary") or entry.get("description") or ""
            key = record_key(doi, link)

            # 昨天已有摘要 -> 直接复用
            if key in prev_has_abs_keys:
                today_records[key] = {**prev_by_key[key], "must_have_abstract": False}
                # 补齐 RSS 日期字段（便于容差逻辑）
                if today_records[key].get("rss_pub_date") in ("", None):
                    today_records[key]["rss_pub_date"] = pub_date
                if not (today_records[key].get("rss_published_str") or "").strip():
                    today_records[key]["rss_published_str"] = published_str
                # pub_date 若为空可用 rss 填一下
                if today_records[key].get("pub_date") is None and pub_date is not None:
                    today_records[key]["pub_date"] = pub_date
                if not (today_records[key].get("published_str") or "").strip() and published_str:
                    today_records[key]["published_str"] = published_str

                if is_sd_special:
                    special_taken += 1
                continue

            abstract = ""
            abstract_source = ""

            if is_cellpress_inpress_any(source_title, link):
                abstract = clean_html_text(desc_html)
                abstract_source = "rss_cell"
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
                "abstract": (abstract or "").strip(),
                "abstract_source": (abstract_source or "").strip(),
                "must_have_abstract": must_have_abstract,
                "pub_date_source": "",
                # RSS 原始日期（容差判定/兜底）
                "rss_pub_date": pub_date,
                "rss_published_str": published_str,
                "wiley_date_diff_days": "",
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

# ================== 摘要补全 + Wiley 日期校验（带容差） ==================

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

    print("\n🔧 API阶段：补全剩余摘要 + Wiley日期校验（Crossref优先；无日期再查PubMed；容差<=3天保留RSS日期）...")
    for r in records:
        doi = (r.get("doi") or "").strip()
        if not doi:
            continue

        # 只对 Wiley 做日期复核
        if is_wiley_record(r.get("source",""), r.get("link",""), doi):
            # 保证 rss_* 存在
            if r.get("rss_pub_date") in ("", None):
                r["rss_pub_date"] = r.get("pub_date")
            if not (r.get("rss_published_str") or "").strip():
                r["rss_published_str"] = r.get("published_str","") or ""

            rss_dt = r.get("rss_pub_date")

            # 1) Crossref
            info = crossref_cache.get(doi)
            if info is None:
                info = query_crossref_abstract(doi)
                crossref_cache[doi] = info
            canonical_dt = info.get("canonical_pub_date")
            canonical_src = "crossref" if canonical_dt else ""

            # 2) PubMed fallback
            if not canonical_dt:
                info2 = pubmed_cache.get(doi)
                if info2 is None:
                    info2 = query_pubmed_abstract(doi)
                    pubmed_cache[doi] = info2
                canonical_dt = info2.get("canonical_pub_date")
                if canonical_dt:
                    canonical_src = "pubmed"

            # 3) 容差判定
            if canonical_dt and rss_dt:
                diff = date_diff_days(canonical_dt, rss_dt)
                r["wiley_date_diff_days"] = str(diff if diff is not None else "")
                if diff is not None and diff <= WILEY_DATE_TOLERANCE_DAYS:
                    # ✅ 差<=3天：完全按 RSS 为准（不改时间；不drop由后面过滤决定）
                    r["pub_date"] = rss_dt
                    r["published_str"] = r.get("rss_published_str") or ""
                    r["pub_date_source"] = f"rss_kept_tol{WILEY_DATE_TOLERANCE_DAYS}"
                else:
                    # ✅ 差>3天：以校验日期为准（后面会用它决定 drop）
                    r["pub_date"] = canonical_dt
                    r["published_str"] = canonical_dt.strftime("%Y-%m-%d %H:%M:%S %Z")
                    r["pub_date_source"] = canonical_src or "verified"
            elif canonical_dt and not rss_dt:
                # RSS无日期：直接用校验日期
                r["wiley_date_diff_days"] = ""
                r["pub_date"] = canonical_dt
                r["published_str"] = canonical_dt.strftime("%Y-%m-%d %H:%M:%S %Z")
                r["pub_date_source"] = canonical_src or "verified"
            else:
                # 无法校验：回退 RSS 日期（不再占位）
                r["wiley_date_diff_days"] = ""
                r["pub_date"] = rss_dt
                r["published_str"] = r.get("rss_published_str") or ""
                r["pub_date_source"] = "rss_fallback_unverified"

            # 顺便用 Crossref 的 abstract（如果有且你还没抽到）
            if not (r.get("abstract") or "").strip():
                abs_txt = (info.get("abstract") or "").strip()
                if abs_txt:
                    r["abstract"] = abs_txt
                    r["abstract_source"] = "crossref"

        # 原有：补全剩余摘要
        if (r.get("abstract") or "").strip():
            continue

        aggressive = is_acs_energy_letters(r.get("source",""), r.get("link",""))
        abs_txt, src = get_abstract_via_apis(doi, aggressive=aggressive)
        if abs_txt:
            r["abstract"] = abs_txt.strip()
            r["abstract_source"] = src

def export_records(today_records: dict):
    if not today_records:
        print("⚠️ 没有记录可导出。")
        return

    df_all = pd.DataFrame(list(today_records.values()))

    # 只允许出现在 CSV 中的列（白名单）
    keep_cols = [
        "title",
        "link",
        "source",
        "published_str",
        "pub_date",
        "doi",
        "abstract",
        "abstract_source",
        "must_have_abstract",
    ]

    # ⚠️ 关键：从 df_all 中“选列”生成新 DataFrame
    df = df_all.loc[:, [c for c in keep_cols if c in df_all.columns]]

    # 防御性补齐缺失列（顺序也在这里统一）
    for c in keep_cols:
        if c not in df.columns:
            df[c] = ""

    # 类型整理
    df["pub_date"] = pd.to_datetime(df["pub_date"], utc=True, errors="coerce")

    for col in ["title", "link", "source", "published_str", "doi", "abstract", "abstract_source"]:
        df[col] = df[col].fillna("").astype(str)

    if "must_have_abstract" in df.columns:
        df["must_have_abstract"] = (
            df["must_have_abstract"].fillna(False).astype(bool).astype(int)
        )

    # 排序
    df = df.sort_values(by="pub_date")

    # 再次强制列顺序 & 列集合
    df = df[keep_cols]

    df.to_csv(TODAY_CSV, index=False, encoding="utf-8-sig")
    print(f"\n✅ 导出：{len(df)} 条 -> {TODAY_CSV}")


# ================== 主流程 ==================

def main():
    print("▶ spider_abstract_1216_wiley_datecheck_keep_rss_tolerance.py 启动")
    print(f"▶ CI={IS_CI}, HEADLESS={HEADLESS}, CHANNEL='{BROWSER_CHANNEL or 'playwright-chromium'}'")
    print(f"▶ 目标日期(UTC)：{sorted(TARGET_DATES)}")
    print(f"▶ Wiley容差：校验日期与RSS日期相差 <= {WILEY_DATE_TOLERANCE_DAYS} 天 => 仍以RSS为准（不改时间、不drop）")
    print(f"▶ 昨日CSV：{YESTERDAY_CSV}")
    print(f"▶ 今日输出：{TODAY_CSV}")
    print("▶ ScienceDirect 特殊源限制 + pub_date写抓取日期的前一天：")
    for k, v in SD_SPECIAL_LIMITS.items():
        print(f"   - {k} -> latest {v}, pub_date={SD_ANCHOR_DATE} (UTC)")
    print("▶ Wiley：Crossref校验；无日期再查PubMed；差>3天才用校验日期并据此决定drop；无法校验则用RSS日期决定去留")

    prev_by_key, prev_has_abs_keys, prev_no_abs_recent3_keys = load_yesterday_cache(YESTERDAY_CSV)

    today_records = collect_rss_records(prev_by_key, prev_has_abs_keys, prev_no_abs_recent3_keys)

    carry_over_prev_with_abstract(today_records, prev_by_key, prev_has_abs_keys)

    records_list = list(today_records.values())
    enrich_with_html_then_api(records_list)
    today_records = {record_key(r.get("doi",""), r.get("link","")): r for r in records_list}

    # ✅ Wiley：最终过滤规则
    # - 若 diff<=3：pub_date 已被保留为 RSS 日期（rss_kept_tol3），只看是否在 TARGET_DATES（基本不会被额外误删）
    # - 若 diff>3：pub_date 已被改为校验日期，若不在 TARGET_DATES => drop
    # - 若无法校验：pub_date 为 RSS 日期，按 RSS 决定去留
    drop_keys = []
    for k, r in today_records.items():
        doi = (r.get("doi") or "").strip()
        if not doi:
            continue
        if is_wiley_record(r.get("source",""), r.get("link",""), doi):
            if not in_target_dates(r.get("pub_date")):
                drop_keys.append(k)

    for k in drop_keys:
        today_records.pop(k, None)

    if drop_keys:
        print(f"🧹 Wiley过滤（基于最终pub_date）：丢弃 {len(drop_keys)} 条（不在目标日期 {sorted(TARGET_DATES)}）")

    retry_records = build_retry_records(prev_by_key, prev_no_abs_recent3_keys, today_records)
    if retry_records:
        enrich_with_html_then_api(retry_records)
        added = 0
        for r in retry_records:
            if (r.get("abstract") or "").strip():
                k = record_key(r.get("doi",""), r.get("link",""))
                doi = (r.get("doi") or "").strip()
                if doi and is_wiley_record(r.get("source",""), r.get("link",""), doi):
                    if not in_target_dates(r.get("pub_date")):
                        continue
                today_records[k] = r
                added += 1
        print(f"✅ 重试成功加入今日：{added} 条（失败的不加入 / Wiley不在目标日期不加入）")

    drop = [k for k, r in today_records.items() if r.get("must_have_abstract") and not (r.get("abstract") or "").strip()]
    for k in drop:
        today_records.pop(k, None)
    if drop:
        print(f"🧹 丢弃 must_have_abstract 且无摘要：{len(drop)} 条")

    export_records(today_records)

if __name__ == "__main__":
    main()
