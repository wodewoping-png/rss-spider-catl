import pandas as pd
import re

# ========== 配置 ==========
INPUT_FILE = "news_with_abstract_2025-12-22.csv"
OUTPUT_FILE = "filtered_news_Qiaochu.csv"

keywords = [
    "CCUS",
    "carbon capture",
    "carbon storage",
    "plastic",
    "deep sea mining",
    "deep sea mine",
    "deep-sea mining",
    "deep-sea mine",
    "CO2",
    "decarbonization",
    "carbon reduction",
    "low-carbon",
    "carbon",
    "carbon neutrality",
    "carbon dioxide"
]

# ========== 读取 ==========
df = pd.read_csv(INPUT_FILE, encoding="utf-8-sig", keep_default_na=False)

# 确保列存在
if "title" not in df.columns:
    raise ValueError("CSV 缺少 title 列")
if "abstract" not in df.columns:
    df["abstract"] = ""  # 没有摘要列就当空

# 统一空值（避免 nan / NaN 干扰）
df["title"] = df["title"].fillna("").astype(str)
df["abstract"] = df["abstract"].fillna("").astype(str)

# 合并用于匹配的全文字段（title + abstract）
df["match_text"] = (df["title"].str.strip() + " " + df["abstract"].str.strip()).str.strip()

# ========== 规则匹配准备 ==========
def tokenize(text: str):
    return re.findall(r"\b\w+\b", text.lower())

phrase_keywords = [kw.lower() for kw in keywords if " " in kw.strip()]
word_keywords = {kw.lower() for kw in keywords if " " not in kw.strip()}

def rule_match(text: str) -> bool:
    t = str(text).lower()

    # 短语必须完整出现
    if any(phrase in t for phrase in phrase_keywords):
        return True

    # 单词匹配：分词后求交集
    tokens = set(tokenize(t))
    if tokens & word_keywords:
        return True

    return False

# 只做规则匹配（title+abstract）
mask_rule = df["match_text"].apply(rule_match)
df_filtered = df[mask_rule].copy()

print(f"规则匹配命中 {len(df_filtered)} 条（无语义筛选）")

# 可选：去掉临时列
df_filtered = df_filtered.drop(columns=["match_text"], errors="ignore")

# ========== 保存 ==========
df_filtered.to_csv(OUTPUT_FILE, index=False, encoding="utf-8-sig")
print(f"✅ 总共筛选 {len(df_filtered)} 条，已保存到 {OUTPUT_FILE}")
