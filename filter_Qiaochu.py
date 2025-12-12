import pandas as pd
import re
from sentence_transformers import SentenceTransformer, util

# 1. 读取新闻数据
df = pd.read_csv("news_week_2025-W49.csv")

# 2. 定义关键词（短语保持原样）
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
    "carbon dioxide"# 单词仍然会走分词匹配
]

# 阈值（0.4~0.7 推荐，越高语义匹配越严格）
threshold = 0.55

# 3. 分词函数
def tokenize(text):
    return re.findall(r"\b\w+\b", text.lower())

# 4. 区分短语和单词
phrase_keywords = [kw.lower() for kw in keywords if " " in kw.strip()]   # 含空格的算短语
word_keywords = {kw.lower() for kw in keywords if " " not in kw.strip()}  # 单词集合

# 5. 子串/分词匹配（第一层）
def rule_match(title: str) -> bool:
    t = str(title).lower()

    # 短语匹配：必须完整出现
    if any(phrase in t for phrase in phrase_keywords):
        return True

    # 单词匹配：分词后看交集
    tokens = set(tokenize(t))
    if tokens & word_keywords:
        return True

    return False

mask_rule = df["title"].apply(rule_match)
df_rule = df[mask_rule]
df_remaining = df[~mask_rule]

print(f"规则匹配命中 {len(df_rule)} 条，进入语义匹配 {len(df_remaining)} 条")

# 6. 语义匹配（第二层）
if not df_remaining.empty:
    model = SentenceTransformer("all-MiniLM-L6-v2")
    titles = df_remaining["title"].astype(str).tolist()
    title_embeddings = model.encode(titles, convert_to_tensor=True)
    query_embeddings = model.encode(keywords, convert_to_tensor=True)

    cos_scores = util.cos_sim(title_embeddings, query_embeddings)
    max_scores = cos_scores.max(dim=1).values.cpu().numpy()

    mask_semantic = max_scores >= threshold
    df_semantic = df_remaining[mask_semantic].copy()
    df_semantic["similarity"] = max_scores[mask_semantic]

    print(f"语义匹配命中 {len(df_semantic)} 条 (阈值={threshold})")
else:
    df_semantic = pd.DataFrame()

# 7. 合并结果
df_filtered = pd.concat([df_rule, df_semantic], ignore_index=True)

# 8. 保存与展示
df_filtered.to_csv("filtered_news_Qiaochuw49.csv", index=False, encoding="utf-8-sig")
print(f"✅ 总共筛选 {len(df_filtered)} 条，已保存到 filtered_news_Qiaochu.csv")
