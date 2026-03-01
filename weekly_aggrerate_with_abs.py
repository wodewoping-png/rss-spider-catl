# weekly_aggregate_with_abs.py
from pathlib import Path
from datetime import datetime, timedelta, timezone
from urllib.parse import urlsplit, urlunsplit
import pandas as pd

OUTPUT_DIR = Path("output")
WEEKLY_DIR = OUTPUT_DIR / "weekly"
WEEKLY_DIR.mkdir(parents=True, exist_ok=True)

# ====== 你要的：滚动 7 天（今天+前6天）======
def daily_file_for(d):
    """
    优先读取新版：news_with_abstract_YYYY-MM-DD.csv
    如果不存在，兼容旧版：news_YYYY-MM-DD.csv
    """
    f_new = OUTPUT_DIR / f"news_with_abstract_{d.strftime('%Y-%m-%d')}.csv"
    if f_new.exists():
        return f_new
    f_old = OUTPUT_DIR / f"news_{d.strftime('%Y-%m-%d')}.csv"
    return f_old

def normalize_link(url: str) -> str:
    """
    link 去重专用：去掉 query/hash，只保留 scheme+host+path
    """
    if not isinstance(url, str) or not url.strip():
        return ""
    parts = urlsplit(url.strip())
    return urlunsplit((parts.scheme, parts.netloc, parts.path, "", ""))

def normalize_pub_date(df: pd.DataFrame) -> pd.DataFrame:
    """
    统一 pub_date 为 UTC aware；若 pub_date 为空，用 published/published_str 兜底解析。
    """
    if "pub_date" in df.columns:
        df["pub_date"] = pd.to_datetime(df["pub_date"], utc=True, errors="coerce")
    else:
        df["pub_date"] = pd.NaT

    if "published" in df.columns:
        missing = df["pub_date"].isna()
        if missing.any():
            pub2 = pd.to_datetime(df.loc[missing, "published"], utc=True, errors="coerce")
            df.loc[missing, "pub_date"] = pub2

    if "published_str" in df.columns:
        missing = df["pub_date"].isna()
        if missing.any():
            pub3 = pd.to_datetime(df.loc[missing, "published_str"], utc=True, errors="coerce")
            df.loc[missing, "pub_date"] = pub3

    return df

def is_nonempty_text(x) -> bool:
    s = "" if x is None else str(x)
    return bool(s.strip())

def pick_nonempty_longer(a: str, b: str) -> str:
    a = "" if a is None else str(a)
    b = "" if b is None else str(b)
    a1 = a.strip()
    b1 = b.strip()
    if not a1 and not b1:
        return ""
    if not a1:
        return b1
    if not b1:
        return a1
    return b1 if len(b1) > len(a1) else a1

def pick_better_abstract(a1: str, a2: str) -> str:
    return pick_nonempty_longer(a1, a2)

def looks_like_generic_title(t: str) -> bool:
    if not t:
        return True
    s = str(t).strip().lower()
    if not s:
        return True
    generic = {"graphical abstract", "table of contents", "toc", "cover image", "no title"}
    if s in generic:
        return True
    if len(s) < 5:
        return True
    return False

def title_quality_score(t: str) -> int:
    if not isinstance(t, str):
        t = "" if t is None else str(t)
    s = t.strip()
    if not s:
        return 0
    score = 10
    if looks_like_generic_title(s):
        score -= 8
    if len(s) < 10:
        score -= 3
    if 20 <= len(s) <= 180:
        score += 3
    if len(s) > 250:
        score -= 2
    return max(score, 0)

def choose_best_title(t1: str, t2: str) -> str:
    t1 = "" if t1 is None else str(t1)
    t2 = "" if t2 is None else str(t2)
    s1 = t1.strip()
    s2 = t2.strip()
    if not s1 and not s2:
        return ""
    if not s1:
        return s2
    if not s2:
        return s1
    q1 = title_quality_score(s1)
    q2 = title_quality_score(s2)
    if q2 > q1:
        return s2
    if q1 > q2:
        return s1
    return s2 if len(s2) > len(s1) else s1

def safe_parse_dt(s: str):
    if not s or not str(s).strip():
        return None
    try:
        return pd.to_datetime(str(s).strip(), utc=True, errors="coerce")
    except Exception:
        return None

def merge_group_rows(g: pd.DataFrame) -> dict:
    """
    同一 link_norm 组内，尽量补齐字段：
    - title/source/doi/last_author: 非空优先；都非空取更长更完整
    - abstract: 更长优先；abstract_source 跟随胜者
    - pub_date: 取最早
    - published/published_str: 取可解析且更早的字符串；否则取更长
    """
    cols = [
        "title", "link", "published", "published_str", "source",
        "pub_date", "doi", "last_author", "last_author_source",
        "abstract", "abstract_source"
    ]
    for c in cols:
        if c not in g.columns:
            g[c] = ""

    pub_dates = pd.to_datetime(g["pub_date"], utc=True, errors="coerce")
    pub_date_min = pub_dates.min() if pub_dates.notna().any() else pd.NaT

    link = ""
    for v in g["link"].tolist():
        if is_nonempty_text(v):
            link = str(v).strip()
            break

    title = ""
    for v in g["title"].tolist():
        title = choose_best_title(title, v)

    source = ""
    for v in g["source"].tolist():
        source = pick_nonempty_longer(source, v)

    doi = ""
    for v in g["doi"].tolist():
        doi = pick_nonempty_longer(doi, v)

    # ✅ last_author + last_author_source：非空优先；都非空取更长；source 跟随胜者
    best_la = ""
    best_la_src = ""
    for la, la_src in zip(g["last_author"].tolist(), g["last_author_source"].tolist()):
        cand = "" if la is None else str(la)
        chosen = pick_nonempty_longer(best_la, cand)
        if chosen != (best_la or "").strip():
            best_la = chosen
            best_la_src = "" if la_src is None else str(la_src).strip()
    last_author = best_la.strip()
    last_author_source = best_la_src.strip() if last_author else ""

    # abstract + abstract_source：更长优先，source 跟随胜者
    best_abs = ""
    best_abs_src = ""
    for a, a_src in zip(g["abstract"].tolist(), g["abstract_source"].tolist()):
        cand = "" if a is None else str(a)
        chosen = pick_better_abstract(best_abs, cand)
        if chosen != (best_abs or "").strip():
            best_abs = chosen
            best_abs_src = "" if a_src is None else str(a_src).strip()
    abstract = best_abs.strip()
    abstract_source = best_abs_src.strip() if abstract else ""

    # published / published_str：取“可解析且更早”的那个；都不可解析就取更长
    candidates = []
    for v in g["published_str"].tolist():
        if is_nonempty_text(v):
            candidates.append(("published_str", str(v).strip()))
    for v in g["published"].tolist():
        if is_nonempty_text(v):
            candidates.append(("published", str(v).strip()))

    best_pub_str = ""
    best_pub_dt = None
    best_pub_kind = ""

    for kind, s in candidates:
        dt = safe_parse_dt(s)
        if dt is not None and pd.notna(dt):
            if best_pub_dt is None or dt < best_pub_dt:
                best_pub_dt = dt
                best_pub_str = s
                best_pub_kind = kind
        else:
            if not best_pub_str:
                best_pub_str = s
                best_pub_kind = kind
            else:
                best_pub_str = pick_nonempty_longer(best_pub_str, s)

    out_published = ""
    out_published_str = ""
    if best_pub_kind == "published_str":
        out_published_str = best_pub_str
    elif best_pub_kind == "published":
        out_published = best_pub_str

    if (not out_published_str) and pd.notna(pub_date_min):
        out_published_str = pd.Timestamp(pub_date_min).to_pydatetime().strftime("%Y-%m-%d %H:%M:%S %Z")

    return {
        "title": title,
        "link": link,
        "published": out_published,
        "published_str": out_published_str,
        "source": source,
        "pub_date": pub_date_min,
        "doi": doi,
        "last_author": last_author,
        "last_author_source": last_author_source,
        "abstract": abstract,
        "abstract_source": abstract_source,
    }

def aggregate_rolling7_dedupe_by_link():
    today_utc = datetime.now(timezone.utc).date()
    start_date = today_utc - timedelta(days=6)
    end_date = today_utc

    files = []
    d = start_date
    while d <= end_date:
        f = daily_file_for(d)
        if f.exists():
            files.append(f)
        d += timedelta(days=1)

    out = WEEKLY_DIR / f"weekly_news_with_abstract_{end_date.strftime('%Y-%m-%d')}.csv"

    # ✅ 输出列：加回 last_author（可选也带 last_author_source）
    empty_cols = [
        "title", "link", "published", "source", "pub_date", "doi",
        "last_author", "last_author_source",
        "abstract", "abstract_source"
    ]

    if not files:
        pd.DataFrame(columns=empty_cols).to_csv(out, index=False, encoding="utf-8-sig")
        print(f"Rolling7 aggregate: 0 records → {out}")
        print(f"Window: {start_date} to {end_date} (UTC dates)")
        return

    dfs = []
    for f in files:
        try:
            df = pd.read_csv(f, encoding="utf-8-sig", keep_default_na=False)
            df = normalize_pub_date(df)

            # 统一列存在 & 类型
            for col in [
                "title", "link", "published", "published_str", "source", "doi",
                "last_author", "last_author_source",
                "abstract", "abstract_source"
            ]:
                if col not in df.columns:
                    df[col] = ""
                else:
                    df[col] = df[col].fillna("").astype(str)

            dfs.append(df)
        except Exception as e:
            print(f"Skip {f} due to error: {e}")

    if not dfs:
        pd.DataFrame(columns=empty_cols).to_csv(out, index=False, encoding="utf-8-sig")
        print(f"Rolling7 aggregate: 0 records → {out}")
        print(f"Window: {start_date} to {end_date} (UTC dates)")
        return

    all_df = pd.concat(dfs, ignore_index=True)

    # ====== 过滤 pub_date 落在滚动7天窗口内 ======
    start_dt = pd.Timestamp(start_date, tz="UTC")
    end_dt_exclusive = pd.Timestamp(end_date + timedelta(days=1), tz="UTC")
    all_df["pub_date"] = pd.to_datetime(all_df["pub_date"], utc=True, errors="coerce")

    in_window = all_df["pub_date"].notna() & (all_df["pub_date"] >= start_dt) & (all_df["pub_date"] < end_dt_exclusive)
    before = len(all_df)
    all_df = all_df.loc[in_window].copy()
    after = len(all_df)
    if before != after:
        print(f"Filtered by pub_date window: dropped {before - after} rows (kept {after})")

    if all_df.empty:
        pd.DataFrame(columns=empty_cols).to_csv(out, index=False, encoding="utf-8-sig")
        print(f"Rolling7 aggregate: 0 records after pub_date filtering → {out}")
        print(f"Window: {start_date} to {end_date} (UTC dates)")
        return

    # ====== 核心：link_norm 分组补齐合并 ======
    all_df["link_norm"] = all_df["link"].apply(normalize_link)
    all_df.loc[all_df["link_norm"] == "", "link_norm"] = (
        "title:" + all_df.loc[all_df["link_norm"] == "", "title"].astype(str).str.strip().str.lower()
    )

    merged_rows = []
    for _, g in all_df.groupby("link_norm", sort=False):
        merged_rows.append(merge_group_rows(g))

    dedup = pd.DataFrame(merged_rows)

    # 保持 weekly 输出不强塞 published_str（但内部用它补过 pub_date）
    dedup = dedup.drop(columns=["published_str"], errors="ignore")

    dedup["pub_date"] = pd.to_datetime(dedup["pub_date"], utc=True, errors="coerce")
    dedup = dedup.sort_values(by="pub_date", ascending=True)

    for c in empty_cols:
        if c not in dedup.columns:
            dedup[c] = ""
    dedup = dedup[empty_cols]

    dedup.to_csv(out, index=False, encoding="utf-8-sig")
    print(f"Rolling7 aggregate: {len(dedup)} records from {len(files)} files → {out}")
    print(f"Window: {start_date} to {end_date} (UTC dates)")

if __name__ == "__main__":
    aggregate_rolling7_dedupe_by_link()
