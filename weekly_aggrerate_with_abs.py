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
    这样 ?af=R、utm、rss 参数不会导致重复。
    """
    if not isinstance(url, str) or not url.strip():
        return ""
    parts = urlsplit(url.strip())
    return urlunsplit((parts.scheme, parts.netloc, parts.path, "", ""))

def normalize_pub_date(df: pd.DataFrame) -> pd.DataFrame:
    """
    统一 pub_date 为 UTC aware；若 pub_date 为空，用 published 兜底解析。
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
    return df

def is_nonempty_text(x) -> bool:
    s = "" if x is None else str(x)
    return bool(s.strip())

def pick_better_abstract(a1: str, a2: str) -> str:
    """
    选更好的 abstract：
    - 非空优先
    - 都非空取更长的（通常更完整）
    """
    a1 = "" if a1 is None else str(a1)
    a2 = "" if a2 is None else str(a2)
    a1s = a1.strip()
    a2s = a2.strip()
    if not a1s and not a2s:
        return ""
    if not a1s:
        return a2s
    if not a2s:
        return a1s
    return a2s if len(a2s) > len(a1s) else a1s

def aggregate_rolling7_dedupe_by_link():
    today_utc = datetime.now(timezone.utc).date()
    start_date = today_utc - timedelta(days=6)
    end_date = today_utc

    # 收集 7 天文件
    files = []
    d = start_date
    while d <= end_date:
        f = daily_file_for(d)
        if f.exists():
            files.append(f)
        d += timedelta(days=1)

    out = WEEKLY_DIR / f"weekly_news_with_abstract_{end_date.strftime('%Y-%m-%d')}.csv"
    empty_cols = ["title", "link", "published", "source", "pub_date", "doi", "abstract", "abstract_source"]

    if not files:
        pd.DataFrame(columns=empty_cols).to_csv(out, index=False, encoding="utf-8-sig")
        print(f"Rolling7 aggregate: 0 records → {out}")
        print(f"Window: {start_date} to {end_date} (UTC dates)")
        return

    dfs = []
    for f in files:
        try:
            # keep_default_na=False：把空字符串保留为 ""，避免 NaN 搞出 nan/空串不统一
            df = pd.read_csv(f, encoding="utf-8-sig", keep_default_na=False)
            df = normalize_pub_date(df)

            # 统一列存在 & 类型
            for col in ["title", "link", "published", "source", "doi", "abstract", "abstract_source"]:
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

    # ====== 核心改动：只用 link 去重（规范化后）======
    all_df["link_norm"] = all_df["link"].apply(normalize_link)

    # 对 link_norm 为空的行：退化为 title 去重 key（否则全挤成一坨）
    # 但你说“只要 link 一样就删”，所以这里只是兜底，不影响 link 有值的主流情况
    all_df.loc[all_df["link_norm"] == "", "link_norm"] = (
        "title:" + all_df.loc[all_df["link_norm"] == "", "title"].astype(str).str.strip().str.lower()
    )

    # 排序：优先让“有摘要”的排在前面；再按 abstract 长度；再按 pub_date 早
    all_df["has_abs"] = all_df["abstract"].apply(is_nonempty_text).astype(int)
    all_df["abs_len"] = all_df["abstract"].astype(str).str.strip().str.len()

    all_df = all_df.sort_values(
        by=["link_norm", "has_abs", "abs_len", "pub_date"],
        ascending=[True, False, False, True],
        kind="mergesort"  # 稳定排序
    )

    # 按 link_norm 去重：保留排序后的第一条（= 有摘要/更长摘要优先）
    dedup = all_df.drop_duplicates(subset=["link_norm"], keep="first").copy()

    # 清理辅助列
    dedup = dedup.drop(columns=["link_norm", "has_abs", "abs_len"], errors="ignore")

    # 最终按 pub_date 排序输出
    dedup["pub_date"] = pd.to_datetime(dedup["pub_date"], utc=True, errors="coerce")
    dedup = dedup.sort_values(by="pub_date", ascending=True)

    dedup.to_csv(out, index=False, encoding="utf-8-sig")
    print(f"Rolling7 aggregate: {len(dedup)} records from {len(files)} files → {out}")
    print(f"Window: {start_date} to {end_date} (UTC dates)")

if __name__ == "__main__":
    aggregate_rolling7_dedupe_by_link()
