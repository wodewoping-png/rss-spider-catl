# weekly_aggregate.py  (rolling 7 days: today + previous 6 days)
from pathlib import Path
from datetime import datetime, timedelta, timezone
from urllib.parse import urlsplit, urlunsplit
import pandas as pd

OUTPUT_DIR = Path("output")
WEEKLY_DIR = OUTPUT_DIR / "weekly"
WEEKLY_DIR.mkdir(parents=True, exist_ok=True)

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
    if not isinstance(url, str) or not url.strip():
        return ""
    parts = urlsplit(url.strip())
    return urlunsplit((parts.scheme, parts.netloc, parts.path, "", ""))

def normalize_pub_date(df: pd.DataFrame) -> pd.DataFrame:
    """
    统一 pub_date 为 UTC aware；若 pub_date 为空，尝试用 published 兜底解析。
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

def pick_better_abstract(a1: str, a2: str) -> str:
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

def build_dedupe_key(row) -> str:
    doi = str(row.get("doi", "") or "").strip().lower()
    if doi:
        return f"doi:{doi}"
    link = normalize_link(str(row.get("link", "") or ""))
    if link:
        return f"url:{link.lower()}"
    title = str(row.get("title", "") or "").strip().lower()
    return f"title:{title}"

def aggregate_rolling_7_days():
    # ====== 窗口：今天(UTC) + 往前6天 ======
    today_utc = datetime.now(timezone.utc).date()
    start_date = today_utc - timedelta(days=6)
    end_date = today_utc  # inclusive by date

    # 收集 7 个 daily 文件
    files = []
    d = start_date
    while d <= end_date:
        f = daily_file_for(d)
        if f.exists():
            files.append(f)
        d += timedelta(days=1)

    out = WEEKLY_DIR / f"news_with_abstract_{end_date.strftime('%Y-%m-%d')}.csv"
    empty_cols = ["title", "link", "published", "source", "pub_date", "doi", "abstract", "abstract_source"]

    if not files:
        pd.DataFrame(columns=empty_cols).to_csv(out, index=False, encoding="utf-8-sig")
        print(f"Rolling7 aggregate: 0 records → {out}")
        print(f"Window: {start_date} to {end_date} (UTC dates)")
        return

    # 读取并合并
    dfs = []
    for f in files:
        try:
            # keep_default_na=False：避免 "" 被读成 NaN，保证 abstract 统一为空串
            df = pd.read_csv(f, encoding="utf-8-sig", keep_default_na=False)
            df = normalize_pub_date(df)

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

    # ====== pub_date 必须在滚动7天窗口内，否则丢弃 ======
    start_dt = pd.Timestamp(start_date, tz="UTC")
    end_dt_exclusive = pd.Timestamp(end_date + timedelta(days=1), tz="UTC")  # [start, end)

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

    # ====== 去重 + abstract 覆盖 ======
    all_df["dedupe_key"] = all_df.apply(build_dedupe_key, axis=1)
    all_df = all_df.sort_values(by="pub_date", ascending=True)

    rows = []
    for _, g in all_df.groupby("dedupe_key", sort=False):
        base = g.iloc[0].copy()

        best_abs = ""
        best_abs_source = str(base.get("abstract_source", "") or "")
        for __, r in g.iterrows():
            cand = r.get("abstract", "")
            new_best = pick_better_abstract(best_abs, cand)
            if new_best != best_abs:
                best_abs = new_best
                best_abs_source = str(r.get("abstract_source", "") or best_abs_source)

        base["abstract"] = best_abs
        base["abstract_source"] = best_abs_source
        rows.append(base)

    dedup = pd.DataFrame(rows).drop(columns=["dedupe_key"], errors="ignore")
    dedup["pub_date"] = pd.to_datetime(dedup["pub_date"], utc=True, errors="coerce")
    dedup = dedup.sort_values(by="pub_date", ascending=True)

    dedup.to_csv(out, index=False, encoding="utf-8-sig")
    print(f"Rolling7 aggregate: {len(dedup)} records from {len(files)} files → {out}")
    print(f"Window: {start_date} to {end_date} (UTC dates)")

if __name__ == "__main__":
    aggregate_rolling_7_days()
