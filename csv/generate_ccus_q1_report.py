import os
from pathlib import Path

import pandas as pd


os.environ.setdefault("MPLCONFIGDIR", "/tmp/matplotlib")

import matplotlib.pyplot as plt


BASE_DIR = Path(__file__).resolve().parent
NEWS_CSV = BASE_DIR / "carbonherald_q1_2026_news.csv"
TAG_COUNTS_CSV = BASE_DIR / "carbonherald_q1_2026_tag_counts.csv"

BAR_PNG = BASE_DIR / "ccus_q1_2026_label_bar.png"
PIE_PNG = BASE_DIR / "ccus_q1_2026_label_pie.png"
MONTHLY_PNG = BASE_DIR / "ccus_q1_2026_monthly_trend.png"
REPORT_MD = BASE_DIR / "ccus_q1_2026_brief_report.md"


def save_label_bar(tag_counts: pd.DataFrame) -> None:
    fig, ax = plt.subplots(figsize=(11, 6))
    ax.bar(tag_counts["label"], tag_counts["news_count"], color="#2F6B7C")
    ax.set_title("CCUS Q1 2026 News by Domain")
    ax.set_xlabel("Domain")
    ax.set_ylabel("News Count")
    ax.tick_params(axis="x", rotation=40)

    for idx, value in enumerate(tag_counts["news_count"]):
        ax.text(idx, value + 1, str(value), ha="center", va="bottom", fontsize=9)

    fig.tight_layout()
    fig.savefig(BAR_PNG, dpi=200)
    plt.close(fig)


def save_label_pie(tag_counts: pd.DataFrame) -> None:
    top_n = 7
    pie_df = tag_counts.head(top_n).copy()
    other_count = int(tag_counts.iloc[top_n:]["news_count"].sum())
    if other_count:
        pie_df.loc[len(pie_df)] = {"label": "Other", "news_count": other_count}

    fig, ax = plt.subplots(figsize=(8, 8))
    ax.pie(
        pie_df["news_count"],
        labels=pie_df["label"],
        autopct="%1.1f%%",
        startangle=90,
        counterclock=False,
    )
    ax.set_title("CCUS Q1 2026 Domain Share")
    fig.tight_layout()
    fig.savefig(PIE_PNG, dpi=200)
    plt.close(fig)


def save_monthly_trend(news: pd.DataFrame) -> pd.Series:
    monthly_counts = news.groupby(news["date"].dt.strftime("%Y-%m")).size()

    fig, ax = plt.subplots(figsize=(8, 4.8))
    ax.plot(monthly_counts.index, monthly_counts.values, marker="o", linewidth=2.2, color="#B85C38")
    ax.set_title("CCUS Q1 2026 Monthly News Trend")
    ax.set_xlabel("Month")
    ax.set_ylabel("News Count")
    ax.grid(axis="y", linestyle="--", alpha=0.35)

    for x, y in zip(monthly_counts.index, monthly_counts.values):
        ax.text(x, y + 1, str(int(y)), ha="center", va="bottom", fontsize=9)

    fig.tight_layout()
    fig.savefig(MONTHLY_PNG, dpi=200)
    plt.close(fig)
    return monthly_counts


def build_report(news: pd.DataFrame, tag_counts: pd.DataFrame, monthly_counts: pd.Series) -> str:
    total_news = len(news)
    total_tag_mentions = int(tag_counts["news_count"].sum())

    top3 = tag_counts.head(3).copy()
    top3["share"] = top3["news_count"] / total_tag_mentions * 100

    source_counts = (
        news["source_categories"]
        .fillna("")
        .str.split("|")
        .explode()
        .replace("", pd.NA)
        .dropna()
        .value_counts()
    )

    monthly_delta = monthly_counts.diff()
    max_month = monthly_counts.idxmax()
    min_month = monthly_counts.idxmin()

    top_labels_by_month = {}
    for month, group in news.groupby(news["date"].dt.strftime("%Y-%m")):
        top_labels_by_month[month] = (
            group["labels"].fillna("").str.split("|").explode().replace("", pd.NA).dropna().value_counts().head(3)
        )

    lines = [
        "# CCUS 2026年一季度趋势简报",
        "",
        "## 1. 数据范围",
        f"- 新闻明细文件：`{NEWS_CSV.name}`",
        f"- 标签统计文件：`{TAG_COUNTS_CSV.name}`",
        "- 统计区间：2026-01-01 至 2026-03-31",
        f"- 新闻总量：{total_news} 篇",
        f"- 标签总提及量：{total_tag_mentions} 次",
        "",
        "## 2. 核心结论",
        f"- 一季度新闻量总体稳定，1月 {int(monthly_counts.iloc[0])} 篇，2月 {int(monthly_counts.iloc[1])} 篇，3月 {int(monthly_counts.iloc[2])} 篇，峰值出现在 {max_month}，低点出现在 {min_month}。",
        f"- 领域关注度高度集中在 `Removal`、`Capture` 和 `Storage`。三者合计 {int(top3['news_count'].sum())} 次，占全部标签提及的 {top3['news_count'].sum() / total_tag_mentions * 100:.1f}%。",
        f"- `Removal` 连续三个月保持第一，分别为 {int(top_labels_by_month['2026-01'].iloc[0])}、{int(top_labels_by_month['2026-02'].iloc[0])}、{int(top_labels_by_month['2026-03'].iloc[0])} 次，说明碳移除仍是季度主线。",
        f"- `Capture` 在 2 月达到季度内月度高点 {int(top_labels_by_month['2026-02'].get('Capture', 0))} 次，随后 3 月回落至 {int(top_labels_by_month['2026-03'].get('Capture', 0))} 次，反映捕集项目新闻在 2 月更集中。",
        f"- 来源类别上，`removal` 相关新闻 {int(source_counts.get('removal', 0))} 条，`capture` 相关新闻 {int(source_counts.get('capture', 0))} 条，前者略高，显示媒体关注仍偏向移除型技术与项目进展。",
        "",
        "## 3. 领域分布",
    ]

    for _, row in tag_counts.iterrows():
        share = row["news_count"] / total_tag_mentions * 100
        lines.append(f"- {row['label']}: {int(row['news_count'])} 次，占比 {share:.1f}%")

    lines.extend(
        [
            "",
            "## 4. 月度走势",
            f"- 2026-01: {int(monthly_counts['2026-01'])} 篇，环比基准月。",
            f"- 2026-02: {int(monthly_counts['2026-02'])} 篇，较 1 月变动 {int(monthly_delta['2026-02']):+d} 篇。",
            f"- 2026-03: {int(monthly_counts['2026-03'])} 篇，较 2 月变动 {int(monthly_delta['2026-03']):+d} 篇。",
            "",
            "各月前三标签：",
        ]
    )

    for month, counts in top_labels_by_month.items():
        summary = ", ".join(f"{label} {int(value)}" for label, value in counts.items())
        lines.append(f"- {month}: {summary}")

    lines.extend(
        [
            "",
            "## 5. 图表文件",
            f"- 条形图：`{BAR_PNG.name}`",
            f"- 饼形图：`{PIE_PNG.name}`",
            f"- 月度趋势图：`{MONTHLY_PNG.name}`",
            "",
            "## 6. 简要判断",
            "- 从季度内节奏看，新闻总量没有剧烈波动，说明 CCUS 议题保持持续曝光，而非由单一事件驱动。",
            "- 从结构看，`Removal` 与 `Capture` 形成双核心，其中移除相关话题更稳定，捕集相关话题更容易受项目签约、融资或政策节点影响。",
            "- `Markets`、`Policy`、`Utilization` 等标签占比次一级，说明产业化与制度建设议题已形成辅助支撑，但尚未超过技术与项目本体的关注度。",
        ]
    )

    return "\n".join(lines) + "\n"


def main() -> None:
    news = pd.read_csv(NEWS_CSV)
    tag_counts = pd.read_csv(TAG_COUNTS_CSV).sort_values("news_count", ascending=False).reset_index(drop=True)
    news["date"] = pd.to_datetime(news["date"])

    save_label_bar(tag_counts)
    save_label_pie(tag_counts)
    monthly_counts = save_monthly_trend(news)

    report = build_report(news, tag_counts, monthly_counts)
    REPORT_MD.write_text(report, encoding="utf-8")

    print(f"Generated: {REPORT_MD.name}")
    print(f"Generated: {BAR_PNG.name}")
    print(f"Generated: {PIE_PNG.name}")
    print(f"Generated: {MONTHLY_PNG.name}")


if __name__ == "__main__":
    main()
