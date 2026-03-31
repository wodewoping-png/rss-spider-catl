#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Merge and classify literature files under ./filtered_news.

Rules:
1. Include all rows from files dated before 2026-03-02.
2. For files dated on/after 2026-03-02, only include sheets named
   "产业降碳" and "CCUS".
3. Merge all selected literature into one CSV, deduplicate by title,
   and keep title / abstract / link.
4. Use SentenceTransformer semantic similarity on title + abstract to
   assign one exclusive category:
   - CO2电化学转化
   - 光或光热转化
   - 热化学转化
   - CO2捕集与分离
   - 其他
5. Priority: first consider the first three conversion categories. If none
   reaches threshold, then consider CO2捕集与分离. Otherwise assign 其他.
"""

from __future__ import annotations

import argparse
import re
from datetime import date
from pathlib import Path
from typing import Dict, Iterable, List, Optional

import pandas as pd
from sentence_transformers import SentenceTransformer, util


CUTOFF_DATE = date(2026, 3, 2)
TARGET_SHEETS_AFTER_CUTOFF = {"产业降碳", "CCUS"}
DEFAULT_MODEL = "sentence-transformers/all-MiniLM-L6-v2"

TITLE_COL_CANDIDATES = ["title", "Title", "标题"]
ABSTRACT_COL_CANDIDATES = ["abstract", "Abstract", "摘要", "abstract_zh"]
LINK_COL_CANDIDATES = ["link", "url", "URL", "链接"]
PUB_DATE_COL_CANDIDATES = ["pub_date", "published", "published_str"]
SOURCE_COL_CANDIDATES = ["source", "来源"]

PRIORITY_CATEGORIES = [
    "CO2电化学转化",
    "光或光热转化",
    "热化学转化",
]
CAPTURE_CATEGORY = "CO2捕集与分离"
OTHER_CATEGORY = "其他"
ALL_OUTPUT_SHEETS = PRIORITY_CATEGORIES + [CAPTURE_CATEGORY, OTHER_CATEGORY]


CATEGORY_PROTOTYPES: Dict[str, List[str]] = {
    "CO2电化学转化": [
        "electrochemical CO2 conversion",
        "electrochemical CO2 reduction",
        "CO2 electroreduction to CO formate methanol methane ethylene ethanol",
        "CO2 electrolyzer electrocatalyst cathode membrane electrode assembly",
        "电催化二氧化碳转化 电还原 电解槽 电解器 甲酸 一氧化碳 甲醇 甲烷 多碳产物",
    ],
    "光或光热转化": [
        "photocatalytic CO2 conversion",
        "photothermal CO2 conversion",
        "solar-driven CO2 reduction photoreduction visible light catalyst",
        "photo-thermal syngas methanol methane from CO2",
        "光催化二氧化碳转化 光热催化 CO2光还原 太阳能驱动 可见光催化",
    ],
    "热化学转化": [
        "thermocatalytic CO2 conversion",
        "thermal catalytic CO2 hydrogenation",
        "reverse water gas shift dry reforming methanol synthesis Fischer Tropsch from CO2",
        "heterogeneous catalysis elevated temperature and pressure for CO2 utilization",
        "热化学二氧化碳转化 热催化 CO2加氢 逆水煤气变换 干重整 甲醇合成",
    ],
    "CO2捕集与分离": [
        "CO2 capture separation adsorption absorption membrane separation direct air capture DAC",
        "carbon capture solvent sorbent amine swing adsorption membrane process",
        "electrochemical CO2 capture from air flue gas ocean water",
        "二氧化碳捕集 分离 吸附 吸收 膜分离 直接空气捕集 电化学捕集",
    ],
}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Merge and classify filtered literature files.")
    parser.add_argument("--input-dir", default="filtered_news", help="Folder containing csv/xlsx files.")
    parser.add_argument(
        "--merged-csv",
        default="filtered_news/all_literature_merged_dedup.csv",
        help="Output CSV path for merged and deduplicated literature.",
    )
    parser.add_argument(
        "--classified-xlsx",
        default="filtered_news/all_literature_classified.xlsx",
        help="Output XLSX path for classified literature.",
    )
    parser.add_argument(
        "--model-name",
        default=DEFAULT_MODEL,
        help="SentenceTransformer model name.",
    )
    parser.add_argument(
        "--threshold",
        type=float,
        default=0.52,
        help="Similarity threshold used for direct classification.",
    )
    return parser.parse_args()


def extract_date_from_name(path: Path) -> Optional[date]:
    match = re.search(r"(\d{4})[-_]?(\d{2})[-_]?(\d{2})", path.stem)
    if not match:
        return None
    year, month, day = map(int, match.groups())
    try:
        return date(year, month, day)
    except ValueError:
        return None


def find_first_column(df: pd.DataFrame, candidates: Iterable[str]) -> Optional[str]:
    lookup = {str(col).strip().lower(): col for col in df.columns}
    for candidate in candidates:
        col = lookup.get(candidate.strip().lower())
        if col is not None:
            return col
    return None


def normalize_text(value: object) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    text = re.sub(r"\s+", " ", text)
    return text


def normalize_title_for_dedup(title: str) -> str:
    title = normalize_text(title).lower()
    title = re.sub(r"\s+", " ", title)
    return title


def normalize_dataframe(df: pd.DataFrame, origin_file: str, origin_sheet: str) -> pd.DataFrame:
    title_col = find_first_column(df, TITLE_COL_CANDIDATES)
    if title_col is None:
        return pd.DataFrame(columns=["title", "abstract", "link", "pub_date", "source", "origin_file", "origin_sheet"])

    abstract_col = find_first_column(df, ABSTRACT_COL_CANDIDATES)
    link_col = find_first_column(df, LINK_COL_CANDIDATES)
    pub_date_col = find_first_column(df, PUB_DATE_COL_CANDIDATES)
    source_col = find_first_column(df, SOURCE_COL_CANDIDATES)

    normalized = pd.DataFrame()
    normalized["title"] = df[title_col].map(normalize_text)
    normalized["abstract"] = df[abstract_col].map(normalize_text) if abstract_col else ""
    normalized["link"] = df[link_col].map(normalize_text) if link_col else ""
    normalized["pub_date"] = df[pub_date_col].map(normalize_text) if pub_date_col else ""
    normalized["source"] = df[source_col].map(normalize_text) if source_col else ""
    normalized["origin_file"] = origin_file
    normalized["origin_sheet"] = origin_sheet

    normalized = normalized[normalized["title"] != ""].copy()
    normalized = normalized.drop_duplicates(subset=["title", "link", "abstract"])
    return normalized


def load_file(path: Path) -> List[pd.DataFrame]:
    file_date = extract_date_from_name(path)

    if path.suffix.lower() == ".csv":
        df = pd.read_csv(path)
        return [normalize_dataframe(df, path.name, "csv")]

    xls = pd.ExcelFile(path)
    selected_sheets = xls.sheet_names
    if file_date is not None and file_date >= CUTOFF_DATE:
        selected_sheets = [sheet for sheet in xls.sheet_names if sheet in TARGET_SHEETS_AFTER_CUTOFF]

    frames: List[pd.DataFrame] = []
    for sheet_name in selected_sheets:
        df = pd.read_excel(path, sheet_name=sheet_name)
        frames.append(normalize_dataframe(df, path.name, sheet_name))
    return frames


def merge_and_dedup(input_dir: Path) -> pd.DataFrame:
    frames: List[pd.DataFrame] = []
    for path in sorted(input_dir.iterdir()):
        if path.suffix.lower() not in {".csv", ".xlsx"}:
            continue
        frames.extend(load_file(path))

    if not frames:
        raise FileNotFoundError(f"No csv/xlsx files found in {input_dir}")

    merged = pd.concat(frames, ignore_index=True)
    merged["dedup_title"] = merged["title"].map(normalize_title_for_dedup)
    merged = merged[merged["dedup_title"] != ""].copy()
    merged["has_abstract"] = merged["abstract"].str.len().fillna(0) > 0
    merged["abstract_len"] = merged["abstract"].str.len().fillna(0)
    merged["link_len"] = merged["link"].str.len().fillna(0)
    deduped = (
        merged.sort_values(
            by=["has_abstract", "abstract_len", "link_len", "pub_date"],
            ascending=[False, False, False, False],
            kind="stable",
        )
        .drop_duplicates(subset=["dedup_title"], keep="first")
        .reset_index(drop=True)
    )
    deduped = deduped.drop(columns=["dedup_title", "has_abstract", "abstract_len", "link_len"], errors="ignore")
    return deduped


def build_text_for_embedding(row: pd.Series) -> str:
    title = normalize_text(row.get("title", ""))
    abstract = normalize_text(row.get("abstract", ""))
    if abstract:
        return f"{title}. {abstract}"
    return title


def classify_records(df: pd.DataFrame, model_name: str, threshold: float) -> pd.DataFrame:
    if df.empty:
        return df.copy()

    model = SentenceTransformer(model_name)
    texts = df.apply(build_text_for_embedding, axis=1).tolist()

    category_names = list(CATEGORY_PROTOTYPES.keys())
    prototype_texts: List[str] = []
    prototype_categories: List[str] = []
    for category, prototypes in CATEGORY_PROTOTYPES.items():
        for text in prototypes:
            prototype_texts.append(text)
            prototype_categories.append(category)

    text_embeddings = model.encode(texts, convert_to_tensor=True, normalize_embeddings=True)
    prototype_embeddings = model.encode(prototype_texts, convert_to_tensor=True, normalize_embeddings=True)
    similarity = util.cos_sim(text_embeddings, prototype_embeddings).cpu().numpy()

    scores_by_category: Dict[str, List[float]] = {name: [] for name in category_names}
    for row_scores in similarity:
        grouped: Dict[str, float] = {name: -1.0 for name in category_names}
        for idx, score in enumerate(row_scores):
            category = prototype_categories[idx]
            grouped[category] = max(grouped[category], float(score))
        for category in category_names:
            scores_by_category[category].append(grouped[category])

    classified = df.copy()
    for category in category_names:
        classified[f"score_{category}"] = scores_by_category[category]

    assigned_labels: List[str] = []
    assigned_scores: List[float] = []
    for _, row in classified.iterrows():
        priority_scores = {category: float(row[f"score_{category}"]) for category in PRIORITY_CATEGORIES}
        capture_score = float(row[f"score_{CAPTURE_CATEGORY}"])

        best_priority_category = max(priority_scores, key=priority_scores.get)
        best_priority_score = priority_scores[best_priority_category]

        if best_priority_score >= threshold:
            assigned_labels.append(best_priority_category)
            assigned_scores.append(best_priority_score)
            continue

        if capture_score >= threshold:
            assigned_labels.append(CAPTURE_CATEGORY)
            assigned_scores.append(capture_score)
            continue

        assigned_labels.append(OTHER_CATEGORY)
        assigned_scores.append(max(best_priority_score, capture_score))

    classified["assigned_category"] = assigned_labels
    classified["assigned_score"] = assigned_scores
    return classified


def save_outputs(df: pd.DataFrame, merged_csv: Path, classified_xlsx: Path) -> None:
    merged_csv.parent.mkdir(parents=True, exist_ok=True)
    classified_xlsx.parent.mkdir(parents=True, exist_ok=True)

    merged_export = df[["title", "abstract", "link"]].copy()
    merged_export.to_csv(merged_csv, index=False, encoding="utf-8-sig")

    export_columns = [
        "title",
        "abstract",
        "link",
        "assigned_category",
        "assigned_score",
        "pub_date",
        "source",
        "origin_file",
        "origin_sheet",
    ] + [f"score_{category}" for category in CATEGORY_PROTOTYPES]

    with pd.ExcelWriter(classified_xlsx, engine="openpyxl") as writer:
        for sheet_name in ALL_OUTPUT_SHEETS:
            sheet_df = df[df["assigned_category"] == sheet_name].copy()
            sheet_df = sheet_df.sort_values(by=["assigned_score", "pub_date"], ascending=[False, False], kind="stable")
            sheet_df[export_columns].to_excel(writer, sheet_name=sheet_name, index=False)


def main() -> None:
    args = parse_args()
    input_dir = Path(args.input_dir)
    if not input_dir.exists():
        raise FileNotFoundError(f"Input directory not found: {input_dir.resolve()}")

    merged = merge_and_dedup(input_dir)
    classified = classify_records(merged, model_name=args.model_name, threshold=args.threshold)
    save_outputs(classified, Path(args.merged_csv), Path(args.classified_xlsx))

    counts = classified["assigned_category"].value_counts().to_dict()
    print(f"Merged records: {len(classified)}")
    print(f"Saved CSV: {Path(args.merged_csv).resolve()}")
    print(f"Saved XLSX: {Path(args.classified_xlsx).resolve()}")
    print("Category counts:")
    for name in ALL_OUTPUT_SHEETS:
        print(f"  {name}: {counts.get(name, 0)}")


if __name__ == "__main__":
    main()
