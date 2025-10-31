import os
import re
import feedparser
import pandas as pd
from html import unescape
from datetime import datetime, timedelta, timezone

# ---------- 正则与工具 ----------

ALT_IMG_RE = re.compile(
    r'<img\b[^>]*\balt\s*=\s*(?P<q>["\'])(?P<alt>.*?)(?P=q)[^>]*>',
    flags=re.IGNORECASE | re.DOTALL
)
TAG_RE = re.compile(r"<[^>]+>")  # 去 HTML 标签

# 从 description 中抓取 "Available online 10 October 2025" 的日期
AVAILABLE_ONLINE_RE = re.compile(
    r"Available\s+online\s+(\d{1,2}\s+[A-Za-z]+\s+\d{4})",
    flags=re.IGNORECASE
)

GENERIC_TITLES = {
    "graphical abstract",
    "table of contents",
    "toc",
    "cover image",
    "no title",
}

MONTHS = {
    "january": 1, "february": 2, "march": 3, "april": 4, "may": 5, "june": 6,
    "july": 7, "august": 8, "september": 9, "october": 10, "november": 11, "december": 12
}

def clean_html_text(s: str) -> str:
    if not s:
        return ""
    s = unescape(s)             # HTML 实体反转义
    s = TAG_RE.sub("", s)       # 去标签
    s = re.sub(r"\s+", " ", s)  # 合并空白
    return s.strip()

def fix_mojibake(s: str) -> str:
    # 处理常见的“鈥�”等引号乱码
    if not s:
        return s
    repl = {
        "鈥�": "'",  # 左/右单引号被错误解码的常见表现
        "鈥": "'",    # 宽松替换
        "â": "'",  # UTF-8→Win1252 典型
        "â": "-",  # en dash
        "â": "-",  # em dash
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
    """
    标题提取优先级：
    1) entry.title（HTML 反转义 + 乱码修正）
    2) 从 content/summary 里 <img alt="..."> 提取
    3) 若 alt 明显更完整且不是占位，则用 alt
    """
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
    """尽量把自由格式日期字符串转 UTC aware；失败返回 None。"""
    if not d:
        return None
    try:
        ts = pd.to_datetime(d, utc=True, errors="raise")
        return ts.to_pydatetime()
    except Exception:
        return None

def parse_available_online_date(description_html: str) -> datetime | None:
    """
    针对 Materials Today / ScienceDirect：
    从 description 中的 "Available online 10 October 2025" 解析日期（UTC 00:00）。
    """
    if not description_html:
        return None
    desc_txt = clean_html_text(description_html)
    m = AVAILABLE_ONLINE_RE.search(desc_txt)
    if not m:
        return None
    date_str = m.group(1)  # e.g., "10 October 2025"
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
        dt = datetime(year, month, day, tzinfo=timezone.utc)
        return dt
    except Exception:
        return parse_date_strict(date_str)

def get_entry_pub_date(entry) -> datetime | None:
    """
    统一获取发布时间（UTC aware）：
    1) published_parsed / updated_parsed
    2) published / updated 字符串
    3) 从 description 中解析 "Available online ..."
    """
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

# ---------- 主流程 ----------

# 读取RSS源列表
with open("feeds1011.txt", "r", encoding="utf-8") as f:
    urls = [line.strip() for line in f if line.strip()]

# 当前日期（UTC）
today = datetime.now(timezone.utc).date()
# 昨天 & 前天（UTC）——【变更1：新增“前天”】
yesterday = today - timedelta(days=1)
day_before = today - timedelta(days=2)
target_dates = {yesterday, day_before}

records = []

for url in urls:
    feed = feedparser.parse(url)
    source_title = feed.feed.get("title", url)

    for entry in feed.entries:
        pub_date = get_entry_pub_date(entry)

        # 过滤：只收“昨天或前天”（UTC 日期）——【变更2：范围从==昨天 改为 in {昨天, 前天}】
        if pub_date and pub_date.date() in target_dates:
            title = get_entry_title(entry)
            link = entry.get("link") or ""
            records.append({
                "title": title,
                "link": link,
                "published": pub_date.strftime("%Y-%m-%d %H:%M:%S %Z"),
                "source": source_title,
                "pub_date": pub_date  # aware datetime (UTC)
            })

# 转 DataFrame
df = pd.DataFrame(records)

# 输出：文件名仍沿用“昨天”的命名不变
os.makedirs("output", exist_ok=True)
file_name = f"output/news_{yesterday.strftime('%Y-%m-%d')}.csv"

if df.empty:
    df.to_csv(file_name, index=False, encoding="utf-8-sig")
    print(f"✅ 抓取完成，共 0 条，已保存到 {file_name}")
    print(f"时间范围：{sorted(target_dates)} (UTC 日期)")
else:
    # 统一 pub_date 为 UTC aware，避免排序类型冲突
    df["pub_date"] = pd.to_datetime(df["pub_date"], utc=True)

    # 标题标准化后去重（保留最早）
    df["title_norm"] = df["title"].fillna("").str.strip()
    df_sorted = df.sort_values(by="pub_date", ascending=True)
    df_dedup = df_sorted.drop_duplicates(subset="title_norm", keep="first") \
                        .drop(columns=["title_norm"])

    df_dedup.to_csv(file_name, index=False, encoding="utf-8-sig")

    print(f"✅ 抓取完成，共 {len(df_dedup)} 条，已保存到 {file_name}")
    print(f"时间范围：{sorted(target_dates)} (UTC 日期)")
