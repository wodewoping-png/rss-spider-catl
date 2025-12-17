# spider_rss_with_abstracts.py
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

# UTC 日期窗口
today_utc = datetime.now(timezone.utc).date()
yesterday_utc = today_utc - timedelta(days=1)
day_before_utc = today_utc - timedelta(days=2)
TARGET_DATES = {yesterday_utc, day_before_utc}

# 昨天 CSV（你昨天生成的那个）
YESTERDAY_CSV = OUTPUT_DIR / f"news_with_abstract_{yesterday_utc.strftime('%Y-%m-%d')}.csv"
# 今天输出 CSV
TODAY_CSV = OUTPUT_DIR / f"news_with_abstract_{today_utc.strftime('%Y-%m-%d')}.csv"

# API 配置
REQ_TIMEOUT = 15
API_SLEEP = 0.5
NCBI_TOOL = "literature_bot"
NCBI_EMAIL = "qiaochuzhang@outlook.com"

# Playwright（只用于 Nature/RSC）
PLAYWRIGHT_HEADLESS = False  # 调试可 False；跑批建议 True

# ================== 通用正则与工具 ==================

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

GENERIC_TITLES = {
    "graphical abstract",
    "table of contents",
    "toc",
    "cover image",
    "no title",
}

ACS_TRASH_TITLES = {
    "issue editorial masthead",
    "issue publication information",
}

# 全局：明显无摘要类型（title关键词过滤）
TITLE_EXCLUDE_KEYWORDS = [
    "editorial",
    "masthead",
    "issue information",
    "cover",
]

# OUP(NSR) RSS description 抽摘要结构
OUP_ABS_RE = re.compile(
    r'boxTitle"\s*>\s*Abstract\s*<\s*/\s*div\s*>\s*(.*?)\s*(?:</span>|</description>|$)',
    flags=re.IGNORECASE | re.DOTALL
)

# ================== 基础函数 ==================

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
    alt = m.group("alt") or ""
    alt = clean_html_text(alt)
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
        if (len(best_alt) >= len(title_from_feed) + 10) and (not looks_generic(best_alt)):
            return best_alt

    return title_from_feed or (alt_candidates[0] if alt_candidates else "")

def parse_date_strict(d: str) -> datetime | None:
    if not d:
        return None
    try:
        ts = pd.to_datetime(d, utc=True, errors="raise")
        return ts.to_pydatetime()
    except Exception:
        return None

def parse_available_online_date(description_html: str) -> datetime | None:
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

def get_entry_pub_date(entry) -> datetime | None:
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
    # drop query/fragment
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
    pub_date = pub_dt.date()
    return (today_utc - pub_date).days <= days

# ================== Publisher 识别 + RSS 抽摘要 ==================

def extract_oup_abstract_from_rss(desc_html: str) -> str:
    if not desc_html:
        return ""
    m = OUP_ABS_RE.search(desc_html)
    if not m:
        return ""
    return clean_html_text(m.group(1))

def is_cellpress_chem_or_joule(source_title: str, link: str) -> bool:
    src_lower = (source_title or "").lower()
    link_lower = (link or "").lower()
    return (
        ("cell.com" in link_lower and ("chem" in link_lower or "joule" in link_lower))
        or ("chem (cell press" in src_lower)
        or ("joule (cell press" in src_lower)
    )

def is_oup_nsr(source_title: str, link: str) -> bool:
    src_lower = (source_title or "").lower()
    link_lower = (link or "").lower()
    return ("national science review" in src_lower) or ("academic.oup.com" in link_lower and "/nsr/" in link_lower)

def is_acs_energy_letters(source_title: str, link: str) -> bool:
    src_lower = (source_title or "").lower()
    link_lower = (link or "").lower()
    return ("acs energy letters" in src_lower) or ("acsenergylett" in link_lower)

# ================== API 摘要：Crossref / S2 / OpenAlex / PubMed ==================

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
    abstract = msg.get("abstract", "") or ""
    if abstract:
        abstract = re.sub(r"<[^>]+>", "", abstract).strip()
    return {"abstract": abstract, "source": "crossref"}

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
    return {"abstract": (data.get("abstract") or ""), "source": "semanticscholar"}

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
    return {"abstract": abstract, "source": "openalex"}

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
    abstract_el_list = article.findall("Abstract/AbstractText")
    if abstract_el_list:
        parts = []
        for el in abstract_el_list:
            text = "".join(el.itertext()).strip()
            if text:
                parts.append(text)
        abstract = "\n".join(parts).strip()

    return {"abstract": abstract, "source": "pubmed"}

def get_abstract_via_apis(doi: str, aggressive: bool = False) -> tuple[str, str]:
    if not doi:
        return "", ""
    print(f"   [API] DOI={doi} 获取摘要 (aggressive={aggressive})")

    if aggressive:
        # ACS Energy Letters：更偏向聚合源
        order = (query_semanticscholar_abstract, query_openalex_abstract, query_crossref_abstract, query_pubmed_abstract)
    else:
        order = (query_crossref_abstract, query_semanticscholar_abstract, query_openalex_abstract, query_pubmed_abstract)

    for fn in order:
        info = fn(doi)
        if info.get("abstract"):
            return info["abstract"], info.get("source", "")
    return "", ""

# ================== Playwright（仅 Nature / RSC） ==================

def extract_nature_abstract(page) -> str:
    try:
        page.wait_for_timeout(800)
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
        page.wait_for_timeout(800)
        el = page.query_selector("div.capsule__text")
        if not el:
            el = page.query_selector("h3.article-abstract__heading + div.capsule__column-wrapper div.capsule__text")
        if not el:
            return ""
        return clean_html_text(el.inner_html())
    except Exception:
        return ""

def get_html_abstract_for_record(page, link: str) -> tuple[str, str]:
    l = (link or "").lower()
    if not l:
        return "", ""
    is_nature = "nature.com" in l
    is_rsc = "rsc.org" in l or "pubs.rsc.org" in l
    if not (is_nature or is_rsc):
        return "", ""

    print(f"   [HTML] Playwright 抓摘要: {link}")
    try:
        time.sleep(random.uniform(0.6, 1.6))
        page.goto(link, wait_until="load", timeout=60_000)
        time.sleep(random.uniform(0.8, 1.8))
    except Exception as e:
        print(f"   ⚠️ 打开页面失败: {e}")
        return "", ""

    if is_nature:
        abs_txt = extract_nature_abstract(page)
        return (abs_txt, "nature_html") if abs_txt else ("", "")
    if is_rsc:
        abs_txt = extract_rsc_abstract(page)
        return (abs_txt, "rsc_html") if abs_txt else ("", "")
    return "", ""

# ================== 新增：读取昨天CSV并构建复用/重试集合 ==================

def load_yesterday_cache(path: Path):
    """
    返回：
      prev_by_key: key -> record(dict)
      prev_has_abs_keys: set
      prev_no_abs_recent3_keys: set（昨天无摘要且3天内，需要今天重试，成功才入今日）
    """
    prev_by_key = {}
    prev_has_abs_keys = set()
    prev_no_abs_recent3_keys = set()

    if not path.exists():
        print(f"ℹ️ 昨天CSV不存在：{path}（首次运行或昨天没产出）")
        return prev_by_key, prev_has_abs_keys, prev_no_abs_recent3_keys

    df = pd.read_csv(path, encoding="utf-8-sig", keep_default_na=False)
    if df.empty:
        return prev_by_key, prev_has_abs_keys, prev_no_abs_recent3_keys

    # pub_date 解析
    if "pub_date" in df.columns:
        df["pub_date"] = pd.to_datetime(df["pub_date"], utc=True, errors="coerce")
    else:
        # 兼容旧列名
        df["pub_date"] = pd.to_datetime(df.get("published_str", ""), utc=True, errors="coerce")

    for _, row in df.iterrows():
        doi = str(row.get("doi", "") or "").strip()
        link = str(row.get("link", "") or "").strip()
        key = record_key(doi, link)

        pub_dt = row.get("pub_date")
        pub_dt = pub_dt.to_pydatetime() if hasattr(pub_dt, "to_pydatetime") else None

        rec = {
            "title": str(row.get("title", "") or ""),
            "link": link,
            "source": str(row.get("source", "") or ""),
            "published_str": str(row.get("published_str", "") or ""),
            "pub_date": pub_dt,
            "doi": doi,
            "abstract": str(row.get("abstract", "") or ""),
            "abstract_source": str(row.get("abstract_source", "") or ""),
        }
        prev_by_key[key] = rec

        has_abs = bool(rec["abstract"].strip())
        if has_abs:
            prev_has_abs_keys.add(key)
        else:
            # 昨天无摘要，且3天内 -> 今天重试；仍无摘要则今天不保存
            if pub_dt and within_days(pub_dt, 3):
                prev_no_abs_recent3_keys.add(key)

    print(f"✅ 读取昨天CSV：{len(prev_by_key)} 条；有摘要 {len(prev_has_abs_keys)}；需重试(3天内无摘要) {len(prev_no_abs_recent3_keys)}")
    return prev_by_key, prev_has_abs_keys, prev_no_abs_recent3_keys

# ================== RSS 收集 ==================

def read_feed_list(path: Path) -> list[str]:
    if not path.exists():
        raise FileNotFoundError(f"RSS 源文件不存在: {path}")
    with path.open("r", encoding="utf-8") as f:
        return [line.strip() for line in f if line.strip()]

def collect_rss_records(prev_by_key, prev_has_abs_keys, prev_no_abs_recent3_keys) -> dict:
    """
    返回 today_records: key -> record(dict)
    这里会做：
      - 遇到昨天已存在且有摘要 -> 直接复用摘要，且不再查
      - NSR/OUP、Cell RSS 直接抽摘要
      - 将“昨天无摘要且3天内”的同key条目标记 must_have_abstract=True（不成功则丢弃）
    """
    urls = read_feed_list(FEED_LIST_FILE)
    today_records = {}

    for url in urls:
        print(f"\n📡 处理 RSS 源: {url}")
        feed = feedparser.parse(url)
        source_title = feed.feed.get("title", url)

        for entry in feed.entries:
            pub_date = get_entry_pub_date(entry)
            if not pub_date or pub_date.date() not in TARGET_DATES:
                continue

            title = get_entry_title(entry)
            if should_drop_by_title(title):
                continue

            link = entry.get("link") or ""
            desc_html = entry.get("summary") or entry.get("description") or ""

            doi = extract_doi_from_url(link)
            key = record_key(doi, link)

            # 如果昨天已有并且有摘要：直接复用，跳过后续查找
            if key in prev_has_abs_keys:
                prev = prev_by_key[key]
                if key not in today_records:
                    today_records[key] = {**prev, "must_have_abstract": False}
                    # 确保 pub_date 字段齐全（昨天csv里可能为空）
                    if not today_records[key].get("pub_date"):
                        today_records[key]["pub_date"] = pub_date
                        today_records[key]["published_str"] = pub_date.strftime("%Y-%m-%d %H:%M:%S %Z")
                continue

            abstract = ""
            abstract_source = ""

            # Cell：RSS description 直接是 abstract
            if is_cellpress_chem_or_joule(source_title, link):
                abstract = clean_html_text(desc_html)
                abstract_source = "rss_cell"

            # OUP(NSR)：RSS description 内有 Abstract box
            elif is_oup_nsr(source_title, link):
                abs_txt = extract_oup_abstract_from_rss(desc_html)
                if abs_txt:
                    abstract = abs_txt
                    abstract_source = "rss_oup"

            must_have_abstract = (key in prev_no_abs_recent3_keys)

            today_records[key] = {
                "title": title,
                "link": link,
                "source": source_title,
                "published_str": pub_date.strftime("%Y-%m-%d %H:%M:%S %Z"),
                "pub_date": pub_date,
                "doi": doi,
                "abstract": abstract,
                "abstract_source": abstract_source,
                "must_have_abstract": must_have_abstract,
            }

    return today_records

# ================== 新增：将昨天有摘要且2天内的直接搬到今天 ==================

def carry_over_prev_with_abstract(today_records: dict, prev_by_key: dict, prev_has_abs_keys: set):
    """
    规则：昨天CSV里 pub_date 在2天内 且 有摘要 -> 直接搬到 today_records
    """
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
    print(f"✅ 从昨天CSV搬运（2天内且有摘要）：{moved} 条")

# ================== 新增：对昨天(3天内无摘要)条目做重试 ==================

def build_retry_records(prev_by_key: dict, prev_no_abs_recent3_keys: set, today_records: dict) -> list[dict]:
    """
    把昨天CSV里 3天内无摘要 的条目加入“重试列表”。
    如果它已经在 today_records 里，说明今天RSS也出现了，就不重复加。
    """
    retry = []
    for key in prev_no_abs_recent3_keys:
        if key in today_records:
            continue
        rec = prev_by_key[key].copy()
        rec["must_have_abstract"] = True
        # 兜底：若 pub_date 缺失则跳过
        if not rec.get("pub_date"):
            continue
        retry.append(rec)
    print(f"✅ 构建重试列表（昨天3天内无摘要 & 今天RSS未出现）：{len(retry)} 条")
    return retry

# ================== 摘要补全（Nature/RSC HTML + APIs） ==================

def enrich_with_html_then_api(records: list[dict]):
    """
    仅对 Nature/RSC 用 Playwright；
    其它走 API；
    ACS Energy Letters aggressive=True。
    """
    # HTML阶段：是否存在需要HTML的记录
    need_html = any(
        ("nature.com" in (r.get("link","").lower()) or "rsc.org" in (r.get("link","").lower()) or "pubs.rsc.org" in (r.get("link","").lower()))
        and not (r.get("abstract") or "").strip()
        for r in records
    )

    if need_html:
        print("\n🔧 HTML阶段：Playwright 抓 Nature / RSC 摘要...")
        with sync_playwright() as p:
            browser = p.chromium.launch(channel="chrome", headless=PLAYWRIGHT_HEADLESS, slow_mo=random.randint(50, 120))
            context = browser.new_context(
                user_agent=(
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/123.0.0.0 Safari/537.36"
                ),
                viewport={"width": random.randint(1200, 1600), "height": random.randint(700, 900)},
            )
            page = context.new_page()

            for r in records:
                if (r.get("abstract") or "").strip():
                    continue
                link = r.get("link", "")
                ll = link.lower()
                if not ("nature.com" in ll or "rsc.org" in ll or "pubs.rsc.org" in ll):
                    continue
                print(f"   [HTML] {r.get('title','')[:80]}")
                abs_txt, src = get_html_abstract_for_record(page, link)
                if abs_txt:
                    r["abstract"] = abs_txt
                    r["abstract_source"] = src

            browser.close()

    # API阶段
    print("\n🔧 API阶段：补全剩余摘要...")
    for r in records:
        if (r.get("abstract") or "").strip():
            continue
        doi = (r.get("doi") or "").strip()
        if not doi:
            continue
        aggressive = is_acs_energy_letters(r.get("source",""), r.get("link",""))
        abs_txt, src = get_abstract_via_apis(doi, aggressive=aggressive)
        if abs_txt:
            r["abstract"] = abs_txt
            r["abstract_source"] = src

# ================== 导出 ==================

def export_records(today_records: dict):
    if not today_records:
        print("⚠️ 没有记录可导出。")
        return

    df = pd.DataFrame(list(today_records.values()))
    df["pub_date"] = pd.to_datetime(df["pub_date"], utc=True, errors="coerce")
    for col in ["abstract", "abstract_source", "doi", "title", "link", "source", "published_str"]:
        if col in df.columns:
            df[col] = df[col].fillna("").astype(str)

    df = df.sort_values(by="pub_date")
    df.to_csv(TODAY_CSV, index=False, encoding="utf-8-sig")
    print(f"\n✅ 导出完成：{len(df)} 条 -> {TODAY_CSV}")

# ================== 主程序 ==================

def main():
    print("▶ RSS + 摘要一体化（含昨日缓存复用/重试）启动...")
    print(f"▶ 目标日期(UTC)：{sorted(TARGET_DATES)}")
    print(f"▶ 昨日CSV：{YESTERDAY_CSV}")
    print(f"▶ 今日输出：{TODAY_CSV}")

    prev_by_key, prev_has_abs_keys, prev_no_abs_recent3_keys = load_yesterday_cache(YESTERDAY_CSV)

    # 1) 先收今天RSS两天窗口（会自动复用昨天已带摘要的同key）
    today_records = collect_rss_records(prev_by_key, prev_has_abs_keys, prev_no_abs_recent3_keys)

    # 2) 再把昨天 2天内且有摘要 的条目搬到今天（防止今天RSS没出现也丢了）
    carry_over_prev_with_abstract(today_records, prev_by_key, prev_has_abs_keys)

    # 3) 对“今天需要处理的记录”做摘要补全（先HTML(Nature/RSC)再API）
    #    注意：这里包含 today_records 里所有条目（含 must_have_abstract 可能为True的）
    records_list = list(today_records.values())
    enrich_with_html_then_api(records_list)

    # 写回
    today_records = {record_key(r.get("doi",""), r.get("link","")): r for r in records_list}

    # 4) 对“昨天3天内无摘要、且今天RSS没出现”的条目做重试
    retry_records = build_retry_records(prev_by_key, prev_no_abs_recent3_keys, today_records)
    if retry_records:
        enrich_with_html_then_api(retry_records)
        # 只把“重试成功拿到摘要”的加进 today_records
        added = 0
        for r in retry_records:
            if (r.get("abstract") or "").strip():
                k = record_key(r.get("doi",""), r.get("link",""))
                today_records[k] = r
                added += 1
        print(f"✅ 重试成功加入今日：{added} 条（失败的不加入）")

    # 5) 最终过滤规则：
    #    - 若 must_have_abstract=True 且仍无摘要 -> 今天不保存
    drop_keys = []
    for k, r in today_records.items():
        if r.get("must_have_abstract") and not (r.get("abstract") or "").strip():
            drop_keys.append(k)
    for k in drop_keys:
        today_records.pop(k, None)
    if drop_keys:
        print(f"🧹 按规则丢弃 must_have_abstract 且无摘要：{len(drop_keys)} 条")

    export_records(today_records)


if __name__ == "__main__":
    main()
