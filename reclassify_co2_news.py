#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Reclassify an existing merged CSV into 11 CO2-related categories.

Input:
- A CSV generated previously, typically with columns such as title / abstract / link.

Output:
- A flat CSV with assigned category and scores.
- An XLSX workbook split by category.
"""

from __future__ import annotations

import argparse
import re
from pathlib import Path
from typing import Dict, Iterable, List, Optional

import pandas as pd
from sentence_transformers import SentenceTransformer, util


DEFAULT_MODEL = "sentence-transformers/all-MiniLM-L6-v2"

TITLE_COL_CANDIDATES = ["title", "Title", "标题"]
ABSTRACT_COL_CANDIDATES = ["abstract", "Abstract", "摘要", "abstract_zh"]
LINK_COL_CANDIDATES = ["link", "url", "URL", "链接"]
PUB_DATE_COL_CANDIDATES = ["pub_date", "published", "published_str", "date"]
SOURCE_COL_CANDIDATES = ["source", "来源"]

CATEGORY_ORDER = [
    "CO₂光转化",
    "CO₂热转化",
    "CO₂电转化",
    "CO₂矿化",
    "CO₂生物转化",
    "CO₂捕集与分离",
    "CO₂运输与封存",
    "系统集成与工艺耦合",
    "能源经济与LCA",
    "政策与产业化",
    "其他",
]

NON_OTHER_CATEGORIES = CATEGORY_ORDER[:-1]
OTHER_CATEGORY = "其他"

DEFAULT_THRESHOLD = 0.42

CATEGORY_PROTOTYPES: Dict[str, List[str]] = {
    "CO₂光转化": [
        "photocatalytic CO2 conversion",
        "photo-driven CO2 reduction",
        "photothermal CO2 conversion solar-driven CO2 hydrogenation",
        "visible-light CO2 reduction photocatalyst",
        "二氧化碳 光催化 光热催化 光驱动转化 光还原 太阳能驱动",
    ],
    "CO₂热转化": [
        "thermocatalytic CO2 conversion",
        "thermal catalytic CO2 hydrogenation reverse water gas shift dry reforming methanol synthesis",
        "high-temperature CO2 utilization thermochemical conversion",
        "二氧化碳 热催化 热化学转化 加氢 逆水煤气变换 干重整 甲醇合成",
    ],
    "CO₂电转化": [
        "electrochemical CO2 conversion",
        "electrochemical CO2 reduction bicarbonate electrolysis CO2 electrolyzer",
        "electrocatalytic CO2 to CO formate methanol methane ethylene ethanol",
        "二氧化碳 电催化转化 电还原 电解 电解槽 电解器 甲酸 一氧化碳 多碳产物",
    ],
    "CO₂矿化": [
        "CO2 mineralization carbonation enhanced weathering mineral curing alkaline materials",
        "mineral sequestration of CO2 in slag cement concrete mine tailings",
        "carbon mineralization carbonate formation for CO2 utilization or storage",
        "二氧化碳 矿化 碳酸盐化 矿物固碳 增强风化 钢渣 水泥 混凝土 尾矿",
    ],
    "CO₂生物转化": [
        "biological CO2 conversion microbial CO2 fixation algae fermentation biosynthesis",
        "enzymatic CO2 conversion bio-based CO2 utilization",
        "photosynthetic microbial electrosynthesis from CO2",
        "二氧化碳 生物转化 微生物固碳 藻类固碳 酶催化 生物合成 发酵",
    ],
    "CO₂捕集与分离": [
        "CO2 capture separation direct air capture DAC flue gas capture",
        "adsorption absorption solvent sorbent membrane separation for CO2",
        "amine capture carbon capture electrolyte CO2 capture materials",
        "二氧化碳 捕集 分离 吸附 吸收 膜分离 直接空气捕集 胺法捕集",
    ],
    "CO₂运输与封存": [
        "CO2 transport and storage CCS CCS infrastructure pipeline ship terminal injection well reservoir",
        "geological storage saline aquifer depleted reservoir offshore storage monitoring",
        "carbon sequestration site characterization storage integrity leakage monitoring",
        "二氧化碳 运输 封存 管道 船运 码头 注入井 地质封存 咸水层 油气藏 监测",
    ],
    "系统集成与工艺耦合": [
        "integrated CO2 capture and conversion process intensification reactor-system coupling",
        "reactive carbon capture integrated electrolyzer heat integration flowsheet optimization",
        "CCUS system integration process coupling hybrid process",
        "二氧化碳 系统集成 工艺耦合 反应分离一体化 过程强化 一体化捕集转化 热集成",
    ],
    "能源经济与LCA": [
        "techno-economic analysis life cycle assessment carbon footprint cost analysis of CO2 technologies",
        "TEA LCA life-cycle costing levelized cost of CO2 capture utilization storage",
        "energy systems analysis of CCUS",
        "二氧化碳 技术经济 生命周期评价 碳足迹 成本分析 能源系统分析 LCA TEA",
    ],
    "政策与产业化": [
        "CCUS policy regulation market mechanism commercialization demonstration deployment",
        "industrialization scale-up standards incentives financing for carbon capture utilization storage",
        "carbon management policy industry roadmap project deployment",
        "二氧化碳 政策 产业化 商业化 示范项目 标准 监管 激励 市场机制 路线图",
    ],
}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Reclassify an existing merged CSV into 11 CO2 categories.")
    parser.add_argument(
        "--input-csv",
        default="filtered_news/all_literature_merged_dedup.csv",
        help="Existing merged CSV path.",
    )
    parser.add_argument(
        "--output-csv",
        default="filtered_news/all_literature_reclassified_11cats.csv",
        help="Output CSV path with assigned category.",
    )
    parser.add_argument(
        "--output-xlsx",
        default="filtered_news/all_literature_reclassified_11cats.xlsx",
        help="Output XLSX path split by category.",
    )
    parser.add_argument(
        "--model-name",
        default=DEFAULT_MODEL,
        help="SentenceTransformer model name.",
    )
    parser.add_argument(
        "--threshold",
        type=float,
        default=DEFAULT_THRESHOLD,
        help="Similarity threshold for assigning a non-'其他' category.",
    )
    return parser.parse_args()


def normalize_text(value: object) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    text = re.sub(r"\s+", " ", text)
    return text


def normalize_title_for_dedup(title: str) -> str:
    title = normalize_text(title).lower()
    return re.sub(r"\s+", " ", title)


def find_first_column(df: pd.DataFrame, candidates: Iterable[str]) -> Optional[str]:
    lookup = {str(col).strip().lower(): col for col in df.columns}
    for candidate in candidates:
        col = lookup.get(candidate.strip().lower())
        if col is not None:
            return col
    return None


def load_existing_csv(path: Path) -> pd.DataFrame:
    df = pd.read_csv(path)

    title_col = find_first_column(df, TITLE_COL_CANDIDATES)
    if title_col is None:
        raise ValueError(f"Cannot find title column in {path}")

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

    normalized = normalized[normalized["title"] != ""].copy()
    normalized["dedup_title"] = normalized["title"].map(normalize_title_for_dedup)
    normalized["has_abstract"] = normalized["abstract"].str.len().fillna(0) > 0
    normalized["abstract_len"] = normalized["abstract"].str.len().fillna(0)
    normalized["link_len"] = normalized["link"].str.len().fillna(0)

    normalized = (
        normalized.sort_values(
            by=["has_abstract", "abstract_len", "link_len", "pub_date"],
            ascending=[False, False, False, False],
            kind="stable",
        )
        .drop_duplicates(subset=["dedup_title"], keep="first")
        .reset_index(drop=True)
    )
    return normalized.drop(columns=["dedup_title", "has_abstract", "abstract_len", "link_len"], errors="ignore")


def build_text_for_embedding(row: pd.Series) -> str:
    title = normalize_text(row.get("title", ""))
    abstract = normalize_text(row.get("abstract", ""))
    if abstract:
        return f"{title}. {abstract}"
    return title


def get_text_blob(row: pd.Series) -> str:
    return f"{normalize_text(row.get('title', ''))} {normalize_text(row.get('abstract', ''))}".lower()


def contains_any(text: str, keywords: Iterable[str]) -> bool:
    for keyword in keywords:
        if re.search(r"[\u4e00-\u9fff]", keyword):
            if keyword in text:
                return True
            continue

        pattern = r"(?<![a-z0-9])" + re.escape(keyword.lower()) + r"(?![a-z0-9])"
        if re.search(pattern, text):
            return True

    return False


def apply_keyword_rules(classified: pd.DataFrame) -> pd.DataFrame:
    route_keywords = {
        "CO₂光转化": [
            "photocatal", "photoreduction", "photo-driven", "photothermal", "solar-driven",
            "visible light", "light-driven", "plasmon", "photochemical", "光催化", "光热", "光还原",
        ],
        "CO₂热转化": [
            "thermocatal", "thermochemical", "reverse water-gas shift", "reverse water gas shift",
            "rwgs", "dry reforming", "co2 hydrogenation", "hydrogenation to methanol",
            "热催化", "热化学", "逆水煤气变换", "干重整", "加氢",
        ],
        "CO₂电转化": [
            "electrochemical", "electroreduction", "electrocatal", "electrolysis", "electrolyzer",
            "electrosynthesis", "bicarbonate electrolysis", "电催化", "电还原", "电解", "电解槽",
        ],
        "CO₂生物转化": [
            "microbial", "biological", "biohybrid", "biosynthesis", "bioconversion", "algae",
            "enzymatic", "fermentation", "微生物", "生物转化", "酶催化", "藻",
        ],
    }
    capture_keywords = [
        "carbon capture", "co2 capture", "direct air capture", "dac", "flue gas", "amine",
        "adsorption", "absorption", "sorbent", "sorbents", "membrane separation",
        "membrane", "capture material", "capture materials", "cof", "mof",
        "metal-organic framework", "metal organic framework", "covalent organic framework",
        "捕集", "分离", "吸附", "吸收", "膜分离",
    ]
    storage_keywords = [
        "pipeline", "terminal", "ship transport", "shipping", "reservoir", "saline aquifer",
        "geological storage", "storage site", "injection well", "subsurface storage",
        "offshore storage", "transport and storage", "sequestration", "封存", "运输", "管道",
        "地质封存", "注入井", "储层", "咸水层",
    ]
    mineral_keywords = [
        "mineralization", "mineralisation", "mineral", "carbonate", "carbonation",
        "enhanced weathering", "concrete curing", "slag", "tailings", "矿化", "碳酸盐化", "增强风化",
    ]
    system_keywords = [
        "integrated capture and conversion", "integrated capture", "reactive carbon capture",
        "process integration", "process coupling", "heat integration", "hybrid process",
        "cascade", "relay", "one-pot", "one pot", "串联", "耦合", "一体化", "过程强化",
    ]
    lca_keywords = [
        "techno-economic", "techno economic", "life cycle assessment", "life-cycle assessment",
        "lca", "tea", "carbon footprint", "levelized cost", "cost analysis", "economic analysis",
        "生命周期", "技术经济", "成本分析", "碳足迹",
    ]
    policy_keywords = [
        "policy", "policies", "regulation", "market", "commercialization", "commercialisation",
        "deployment", "demonstration", "roadmap", "subsidy", "tax", "taxes", "industry",
        "industrialization", "industrialisation", "certification", "政策", "产业化", "商业化",
        "示范", "路线图", "补贴", "税",
    ]
    synthesis_keywords = [
        "synthesis", "synthesize", "synthesizing", "synthetic", "electrosynthesis", "organic synthesis",
        "polymer", "polyols", "methanol", "ethanol", "ethylene", "acetate", "formate", "syngas",
        "aldehyde", "urea", "chemical production", "value-added chemicals", "聚合物", "有机合成", "化学品",
    ]
    conversion_keywords = [
        "reduction", "conversion", "utilization", "utilisation", "production", "product", "products",
        "electroreduction", "photoreduction", "hydrogenation", "faradaic efficiency", "selectivity",
        "利用", "转化", "还原", "产物",
    ]
    co2_keywords = [
        "co2", "carbon dioxide", "二氧化碳",
    ]

    adjusted = classified.copy()
    adjusted["rule_category"] = ""
    adjusted["rule_reason"] = ""

    for idx, row in adjusted.iterrows():
        text = get_text_blob(row)
        has_co2 = contains_any(text, co2_keywords)
        if not has_co2:
            continue

        best_category = max(NON_OTHER_CATEGORIES, key=lambda category: float(row[f"score_{category}"]))
        best_score = float(row[f"score_{best_category}"])

        has_mof_cof = contains_any(
            text,
            [
                "mof", "cof", "metal-organic framework", "metal organic framework",
                "metal-organic frameworks", "metal organic frameworks",
                "covalent organic framework", "covalent organic frameworks", "框架材料",
            ],
        )

        if contains_any(text, storage_keywords):
            adjusted.at[idx, "assigned_category"] = "CO₂运输与封存"
            adjusted.at[idx, "assigned_score"] = max(best_score, float(row["score_CO₂运输与封存"]))
            adjusted.at[idx, "rule_category"] = "CO₂运输与封存"
            adjusted.at[idx, "rule_reason"] = "storage_transport_keywords"
            continue

        matched_route = None
        for category, keywords in route_keywords.items():
            if contains_any(text, keywords):
                matched_route = category
                break

        if has_mof_cof and matched_route == "CO₂光转化":
            adjusted.at[idx, "assigned_category"] = "CO₂光转化"
            adjusted.at[idx, "assigned_score"] = max(best_score, float(row["score_CO₂光转化"]))
            adjusted.at[idx, "rule_category"] = "CO₂光转化"
            adjusted.at[idx, "rule_reason"] = "mof_cof_photo_route_keywords"
            continue

        if has_mof_cof and matched_route == "CO₂电转化":
            adjusted.at[idx, "assigned_category"] = "CO₂电转化"
            adjusted.at[idx, "assigned_score"] = max(best_score, float(row["score_CO₂电转化"]))
            adjusted.at[idx, "rule_category"] = "CO₂电转化"
            adjusted.at[idx, "rule_reason"] = "mof_cof_electro_route_keywords"
            continue

        if has_mof_cof and matched_route == "CO₂热转化":
            adjusted.at[idx, "assigned_category"] = "CO₂热转化"
            adjusted.at[idx, "assigned_score"] = max(best_score, float(row["score_CO₂热转化"]))
            adjusted.at[idx, "rule_category"] = "CO₂热转化"
            adjusted.at[idx, "rule_reason"] = "mof_cof_thermal_route_keywords"
            continue

        if matched_route and contains_any(text, conversion_keywords):
            adjusted.at[idx, "assigned_category"] = matched_route
            adjusted.at[idx, "assigned_score"] = max(best_score, float(row[f"score_{matched_route}"]))
            adjusted.at[idx, "rule_category"] = matched_route
            adjusted.at[idx, "rule_reason"] = "conversion_route_priority"
            continue

        if contains_any(text, capture_keywords):
            adjusted.at[idx, "assigned_category"] = "CO₂捕集与分离"
            adjusted.at[idx, "assigned_score"] = max(best_score, float(row["score_CO₂捕集与分离"]))
            adjusted.at[idx, "rule_category"] = "CO₂捕集与分离"
            adjusted.at[idx, "rule_reason"] = "capture_separation_keywords"
            continue

        if contains_any(text, lca_keywords):
            adjusted.at[idx, "assigned_category"] = "能源经济与LCA"
            adjusted.at[idx, "assigned_score"] = max(best_score, float(row["score_能源经济与LCA"]))
            adjusted.at[idx, "rule_category"] = "能源经济与LCA"
            adjusted.at[idx, "rule_reason"] = "lca_keywords"
            continue

        if contains_any(text, policy_keywords):
            adjusted.at[idx, "assigned_category"] = "政策与产业化"
            adjusted.at[idx, "assigned_score"] = max(best_score, float(row["score_政策与产业化"]))
            adjusted.at[idx, "rule_category"] = "政策与产业化"
            adjusted.at[idx, "rule_reason"] = "policy_keywords"
            continue

        if contains_any(text, mineral_keywords):
            adjusted.at[idx, "assigned_category"] = "CO₂矿化"
            adjusted.at[idx, "assigned_score"] = max(best_score, float(row["score_CO₂矿化"]))
            adjusted.at[idx, "rule_category"] = "CO₂矿化"
            adjusted.at[idx, "rule_reason"] = "mineral_keywords"
            continue

        if matched_route and contains_any(text, synthesis_keywords):
            adjusted.at[idx, "assigned_category"] = matched_route
            adjusted.at[idx, "assigned_score"] = max(best_score, float(row[f"score_{matched_route}"]))
            adjusted.at[idx, "rule_category"] = matched_route
            adjusted.at[idx, "rule_reason"] = "organic_synthesis_route_keywords"
            continue

        if matched_route:
            adjusted.at[idx, "assigned_category"] = matched_route
            adjusted.at[idx, "assigned_score"] = max(best_score, float(row[f"score_{matched_route}"]))
            adjusted.at[idx, "rule_category"] = matched_route
            adjusted.at[idx, "rule_reason"] = "route_keywords"
            continue

        if contains_any(text, system_keywords):
            adjusted.at[idx, "assigned_category"] = "系统集成与工艺耦合"
            adjusted.at[idx, "assigned_score"] = max(best_score, float(row["score_系统集成与工艺耦合"]))
            adjusted.at[idx, "rule_category"] = "系统集成与工艺耦合"
            adjusted.at[idx, "rule_reason"] = "system_keywords"

    return adjusted


def classify_records(df: pd.DataFrame, model_name: str, threshold: float) -> pd.DataFrame:
    if df.empty:
        classified = df.copy()
        classified["assigned_category"] = []
        classified["assigned_score"] = []
        return classified

    model = SentenceTransformer(model_name)
    texts = df.apply(build_text_for_embedding, axis=1).tolist()

    prototype_texts: List[str] = []
    prototype_categories: List[str] = []
    for category, prototypes in CATEGORY_PROTOTYPES.items():
        for text in prototypes:
            prototype_texts.append(text)
            prototype_categories.append(category)

    text_embeddings = model.encode(texts, convert_to_tensor=True, normalize_embeddings=True)
    prototype_embeddings = model.encode(prototype_texts, convert_to_tensor=True, normalize_embeddings=True)
    similarity = util.cos_sim(text_embeddings, prototype_embeddings).cpu().numpy()

    scores_by_category: Dict[str, List[float]] = {name: [] for name in NON_OTHER_CATEGORIES}
    for row_scores in similarity:
        grouped: Dict[str, float] = {name: -1.0 for name in NON_OTHER_CATEGORIES}
        for idx, score in enumerate(row_scores):
            category = prototype_categories[idx]
            grouped[category] = max(grouped[category], float(score))
        for category in NON_OTHER_CATEGORIES:
            scores_by_category[category].append(grouped[category])

    classified = df.copy()
    for category in NON_OTHER_CATEGORIES:
        classified[f"score_{category}"] = scores_by_category[category]

    assigned_labels: List[str] = []
    assigned_scores: List[float] = []
    for _, row in classified.iterrows():
        best_category = max(NON_OTHER_CATEGORIES, key=lambda category: float(row[f"score_{category}"]))
        best_score = float(row[f"score_{best_category}"])

        if best_score >= threshold:
            assigned_labels.append(best_category)
            assigned_scores.append(best_score)
        else:
            assigned_labels.append(OTHER_CATEGORY)
            assigned_scores.append(best_score)

    classified["assigned_category"] = assigned_labels
    classified["assigned_score"] = assigned_scores
    return apply_keyword_rules(classified)


def save_outputs(df: pd.DataFrame, output_csv: Path, output_xlsx: Path) -> None:
    output_csv.parent.mkdir(parents=True, exist_ok=True)
    output_xlsx.parent.mkdir(parents=True, exist_ok=True)

    score_columns = [f"score_{category}" for category in NON_OTHER_CATEGORIES]
    export_columns = [
        "title",
        "abstract",
        "link",
        "assigned_category",
        "assigned_score",
        "rule_category",
        "rule_reason",
        "pub_date",
        "source",
    ] + score_columns

    df[export_columns].to_csv(output_csv, index=False, encoding="utf-8-sig")

    with pd.ExcelWriter(output_xlsx, engine="openpyxl") as writer:
        for sheet_name in CATEGORY_ORDER:
            sheet_df = df[df["assigned_category"] == sheet_name].copy()
            sheet_df = sheet_df.sort_values(by=["assigned_score", "pub_date"], ascending=[False, False], kind="stable")
            sheet_df[export_columns].to_excel(writer, sheet_name=sheet_name, index=False)


def main() -> None:
    args = parse_args()
    input_csv = Path(args.input_csv)
    if not input_csv.exists():
        raise FileNotFoundError(f"Input CSV not found: {input_csv.resolve()}")

    records = load_existing_csv(input_csv)
    classified = classify_records(records, model_name=args.model_name, threshold=args.threshold)
    save_outputs(classified, Path(args.output_csv), Path(args.output_xlsx))

    counts = classified["assigned_category"].value_counts().to_dict()
    print(f"Loaded records: {len(records)}")
    print(f"Saved CSV: {Path(args.output_csv).resolve()}")
    print(f"Saved XLSX: {Path(args.output_xlsx).resolve()}")
    print("Category counts:")
    for name in CATEGORY_ORDER:
        print(f"  {name}: {counts.get(name, 0)}")


if __name__ == "__main__":
    main()
