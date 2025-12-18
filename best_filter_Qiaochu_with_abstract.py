import pandas as pd
import re
from sentence_transformers import SentenceTransformer, util

# ========== 配置 ==========
INPUT_FILE = "news_with_abstract_2025-12-18.csv"
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

threshold = 0.55  # 0.4~0.7 推荐，越高越严格

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

# 第一层：对 title+abstract 做规则匹配
mask_rule = df["match_text"].apply(rule_match)
df_rule = df[mask_rule].copy()
df_remaining = df[~mask_rule].copy()

print(f"规则匹配命中 {len(df_rule)} 条，进入语义匹配 {len(df_remaining)} 条")

# ========== 语义匹配（第二层） ==========
df_semantic = pd.DataFrame()
if not df_remaining.empty:
    model = SentenceTransformer("all-MiniLM-L6-v2")

    # 用 title+abstract 作为语义输入（比只用 title 更强）
    texts = df_remaining["match_text"].astype(str).tolist()

    text_embeddings = model.encode(texts, convert_to_tensor=True, show_progress_bar=True)
    query_embeddings = model.encode(keywords, convert_to_tensor=True, show_progress_bar=False)

    cos_scores = util.cos_sim(text_embeddings, query_embeddings)
    max_scores = cos_scores.max(dim=1).values.cpu().numpy()

    mask_semantic = max_scores >= threshold
    df_semantic = df_remaining.loc[mask_semantic].copy()
    df_semantic["similarity"] = max_scores[mask_semantic]

    print(f"语义匹配命中 {len(df_semantic)} 条 (阈值={threshold})")

# ========== 合并结果 ==========
df_filtered = pd.concat([df_rule, df_semantic], ignore_index=True)

# 可选：去掉临时列
df_filtered = df_filtered.drop(columns=["match_text"], errors="ignore")

# ========== 保存 ==========
df_filtered.to_csv(OUTPUT_FILE, index=False, encoding="utf-8-sig")
print(f"✅ 总共筛选 {len(df_filtered)} 条，已保存到 {OUTPUT_FILE}")
