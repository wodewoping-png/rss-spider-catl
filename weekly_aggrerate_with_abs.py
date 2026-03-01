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
    统一 pub_date 为 UTC aware；若 pub_date 为空，用 published/published_str 兜底解析。
    """
    if "pub_date" in df.columns:
        df["pub_date"] = pd.to_datetime(df["pub_date"], utc=True, errors="coerce")
    else:
        df["pub_date"] = pd.NaT

    # 兼容旧列 published
    if "published" in df.columns:
        missing = df["pub_date"].isna()
        if missing.any():
            pub2 = pd.to_datetime(df.loc[missing, "published"], utc=True, errors="coerce")
            df.loc[missing, "pub_date"] = pub2

    # 兼容新版 published_str
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
    """
    通用：非空优先；都非空取更长的（更完整）
    """
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
    """
    选更好的 abstract：
    - 非空优先
    - 都非空取更长的（通常更完整）
    """
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
    """
    更像“真标题”的分数更高
    """
    if not isinstance(t, str):
        t = "" if t is None else str(t)
    s = t.strip()
    if not s:
        return 0
    score = 10
    if looks_like_generic_title(s):
        score -= 8
    # 太短扣分
    if len(s) < 10:
        score -= 3
    # 适中长度加分（通常更像论文标题）
    if 20 <= len(s) <= 180:
        score += 3
    # 超长也扣点（可能是摘要/乱串）
    if len(s) > 250:
        score -= 2
    return max(score, 0)

def choose_best_title(t1: str, t2: str) -> str:
    """
    标题：先比质量分，再比长度，再比非空
    """
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
    # 分数相同，取更长更完整
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
    把同一个 link_norm 组里的多条记录“补齐合并”成一条：
    - title/source/doi: 非空优先+更好更长
    - abstract: 更长优先，并让 abstract_source 跟随最终 abstract
    - pub_date: 取最早（更稳）
    - published/published_str: 取最能解析且更早的字符串（尽量保留）
    - link: 取非空（通常一样）
    """
    # 保证列存在
    cols = ["title", "link", "published", "published_str", "source", "pub_date", "doi", "abstract", "abstract_source"]
    for c in cols:
        if c not in g.columns:
            g[c] = ""

    # 先准备 pub_date（Timestamp）
    pub_dates = pd.to_datetime(g["pub_date"], utc=True, errors="coerce")
    pub_date_min = pub_dates.min() if pub_dates.notna().any() else pd.NaT

    # 选 link（非空优先）
    link = ""
    for v in g["link"].tolist():
        if is_nonempty_text(v):
            link = str(v).strip()
            break

    # 选 title（在组里择优）
    title = ""
    for v in g["title"].tolist():
        title = choose_best_title(title, v)

    # source：通常相同，取更长更具体的
    source = ""
    for v in g["source"].tolist():
        source = pick_nonempty_longer(source, v)

    # doi：非空优先（一般不会冲突）
    doi = ""
    for v in g["doi"].tolist():
        doi = pick_nonempty_longer(doi, v)

    # abstract + abstract_source：用“更好摘要”规则，source 跟随胜者
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

    # published / published_str：取“可解析且更早”的那个；都不可解析就取更长非空
    published = ""
    published_dt = None

    # 优先考虑 published_str，其次 published（但输出里两个都保留兼容）
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
            # 不可解析：留作兜底
            if not best_pub_str:
                best_pub_str = s
                best_pub_kind = kind
            else:
                # 都不可解析：取更长
                best_pub_str = pick_nonempty_longer(best_pub_str, s)

    # 写回 published / published_str：尽量保持“字段语义”
    out_published = ""
    out_published_str = ""
    if best_pub_kind == "published_str":
        out_published_str = best_pub_str
    elif best_pub_kind == "published":
        out_published = best_pub_str
    else:
        # 没有候选
        out_published = ""
        out_published_str = ""

    # 如果有 pub_date_min 但 published_str 空，补一个（不改变你的输出列集合，只是补值）
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
        "abstract": abstract,
        "abstract_source": abstract_source,
    }

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
            df = pd.read_csv(f, encoding="utf-8-sig", keep_default_na=False)
            df = normalize_pub_date(df)

            # 统一列存在 & 类型
            for col in ["title", "link", "published", "published_str", "source", "doi", "abstract", "abstract_source"]:
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

    # ====== 只用 link 去重（规范化后）======
    all_df["link_norm"] = all_df["link"].apply(normalize_link)

    # link_norm 为空：退化为 title key（兜底）
    all_df.loc[all_df["link_norm"] == "", "link_norm"] = (
        "title:" + all_df.loc[all_df["link_norm"] == "", "title"].astype(str).str.strip().str.lower()
    )

    # ====== ✅ 核心改动：按 link_norm 分组“补齐合并” ======
    merged_rows = []
    for key, g in all_df.groupby("link_norm", sort=False):
        merged_rows.append(merge_group_rows(g))

    dedup = pd.DataFrame(merged_rows)

    # 输出列兼容：你原来 weekly 输出里没有 published_str，这里也不强塞
    # 但为了不丢信息：如果你愿意，也可以把 published_str 加入 empty_cols/输出列
    # 目前：保留 published 这个旧列，若 daily 是新格式，published_str 会用于补 pub_date，但不会输出成新列
    if "published_str" in dedup.columns:
        # 不输出 published_str（保持你当前 weekly 输出列集合不变）
        dedup = dedup.drop(columns=["published_str"], errors="ignore")

    # 最终按 pub_date 排序输出
    dedup["pub_date"] = pd.to_datetime(dedup["pub_date"], utc=True, errors="coerce")
    dedup = dedup.sort_values(by="pub_date", ascending=True)

    # 确保输出列齐全
    for c in empty_cols:
        if c not in dedup.columns:
            dedup[c] = ""
    dedup = dedup[empty_cols]

    dedup.to_csv(out, index=False, encoding="utf-8-sig")
    print(f"Rolling7 aggregate: {len(dedup)} records from {len(files)} files → {out}")
    print(f"Window: {start_date} to {end_date} (UTC dates)")

if __name__ == "__main__":
    aggregate_rolling7_dedupe_by_link()
