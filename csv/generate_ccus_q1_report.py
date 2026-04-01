from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Alignment, Font, PatternFill


BASE_DIR = Path(__file__).resolve().parent
NEWS_CSV = BASE_DIR / "carbonherald_q1_2026_news.csv"
TAG_COUNTS_CSV = BASE_DIR / "carbonherald_q1_2026_tag_counts.csv"
REPORT_XLSX = BASE_DIR / "ccus_q1_2026_中文分析报告.xlsx"
MIN_NEWS_COUNT = 15

LABEL_ZH = {
    "Removal": "碳移除",
    "Capture": "碳捕集",
    "Storage": "碳封存",
    "Biomass": "生物质",
    "Markets": "碳市场",
    "Direct Air Capture": "直接空气捕集",
    "Mineralization": "矿化",
    "Policy": "政策",
    "Forests": "森林碳汇",
    "Farming": "农业碳汇",
    "Utilization": "碳利用",
    "BECCS": "生物质能碳捕集与封存",
    "Ocean": "海洋碳汇",
}

SOURCE_CATEGORY_ZH = {
    "removal": "碳移除",
    "capture": "碳捕集",
    "storage": "碳封存",
    "utilization": "碳利用",
    "policy": "政策",
    "markets": "碳市场",
}

HEADER_FILL = PatternFill("solid", fgColor="D9EAF7")
TITLE_FILL = PatternFill("solid", fgColor="1F4E78")
SECTION_FILL = PatternFill("solid", fgColor="DCE6F1")
TITLE_FONT = Font(name="Microsoft YaHei", size=14, bold=True, color="FFFFFF")
SECTION_FONT = Font(name="Microsoft YaHei", size=11, bold=True)
BODY_FONT = Font(name="Microsoft YaHei", size=10)


def to_zh_label(label: str) -> str:
    return LABEL_ZH.get(label, label)


def prepare_data() -> tuple[pd.DataFrame, pd.DataFrame, pd.Series, dict[str, pd.Series], pd.Series, pd.DataFrame]:
    news = pd.read_csv(NEWS_CSV)
    news["date"] = pd.to_datetime(news["date"])

    tag_counts = (
        pd.read_csv(TAG_COUNTS_CSV)
        .sort_values("news_count", ascending=False)
        .query("news_count >= @MIN_NEWS_COUNT")
        .reset_index(drop=True)
        .copy()
    )
    tag_counts["label_zh"] = tag_counts["label"].map(to_zh_label)
    tag_counts["share_pct"] = tag_counts["news_count"] / tag_counts["news_count"].sum()

    monthly_counts = news.groupby(news["date"].dt.strftime("%Y-%m")).size().sort_index()

    top_labels_by_month = {}
    for month, group in news.groupby(news["date"].dt.strftime("%Y-%m")):
        counts = group["labels"].fillna("").str.split("|").explode().replace("", pd.NA).dropna().value_counts().head(3)
        counts.index = counts.index.map(to_zh_label)
        top_labels_by_month[month] = counts

    source_counts = (
        news["source_categories"]
        .fillna("")
        .str.split("|")
        .explode()
        .replace("", pd.NA)
        .dropna()
        .value_counts()
    )

    monthly_label_dist = (
        news.assign(month=news["date"].dt.strftime("%Y-%m"))
        .assign(label=news["labels"].fillna("").str.split("|"))
        .explode("label")
        .replace({"label": {"": pd.NA}})
        .dropna(subset=["label"])
        .assign(label_zh=lambda df: df["label"].map(to_zh_label))
        .groupby(["label_zh", "month"])
        .size()
        .unstack(fill_value=0)
        .reindex(tag_counts["label_zh"].tolist(), fill_value=0)
    )
    monthly_label_dist = monthly_label_dist.reindex(sorted(monthly_counts.index), axis=1, fill_value=0)

    return news, tag_counts, monthly_counts, top_labels_by_month, source_counts, monthly_label_dist


def build_summary_lines(
    news: pd.DataFrame,
    tag_counts: pd.DataFrame,
    monthly_counts: pd.Series,
    top_labels_by_month: dict[str, pd.Series],
    source_counts: pd.Series,
) -> list[str]:
    total_news = len(news)
    total_tag_mentions = int(tag_counts["news_count"].sum())
    top3 = tag_counts.head(3)
    max_month = monthly_counts.idxmax()
    min_month = monthly_counts.idxmin()
    monthly_delta = monthly_counts.diff()

    removal_by_month = [
        int(top_labels_by_month.get(month, pd.Series(dtype="int64")).get("碳移除", 0))
        for month in ["2026-01", "2026-02", "2026-03"]
    ]

    lines = [
        "一、数据范围",
        f"新闻明细文件：{NEWS_CSV.name}",
        f"标签统计文件：{TAG_COUNTS_CSV.name}",
        "统计区间：2026-01-01 至 2026-03-31",
        f"新闻总量：{total_news} 篇",
        f"纳入统计的标签阈值：不少于 {MIN_NEWS_COUNT} 次",
        f"纳入统计的标签提及量：{total_tag_mentions} 次",
        "",
        "二、核心结论",
        (
            f"一季度新闻量总体稳定，1月 {int(monthly_counts['2026-01'])} 篇，"
            f"2月 {int(monthly_counts['2026-02'])} 篇，3月 {int(monthly_counts['2026-03'])} 篇，"
            f"峰值出现在 {max_month}，低点出现在 {min_month}。"
        ),
        (
            f"领域关注度高度集中在“{top3.iloc[0]['label_zh']}”、“{top3.iloc[1]['label_zh']}”和“{top3.iloc[2]['label_zh']}”。"
            f"三者合计 {int(top3['news_count'].sum())} 次，占全部标签提及的 "
            f"{top3['news_count'].sum() / total_tag_mentions * 100:.1f}%。"
        ),
        (
            f"“碳移除”连续三个月保持第一，分别为 {removal_by_month[0]}、{removal_by_month[1]}、"
            f"{removal_by_month[2]} 次，说明碳移除仍是季度主线。"
        ),
        (
            f"“碳捕集”在 2026-02 达到季度内月度高点 "
            f"{int(top_labels_by_month['2026-02'].get('碳捕集', 0))} 次，"
            f"随后 2026-03 回落至 {int(top_labels_by_month['2026-03'].get('碳捕集', 0))} 次。"
        ),
        (
            f"来源类别上，“{SOURCE_CATEGORY_ZH.get('removal', 'removal')}”相关新闻 "
            f"{int(source_counts.get('removal', 0))} 条，"
            f"“{SOURCE_CATEGORY_ZH.get('capture', 'capture')}”相关新闻 "
            f"{int(source_counts.get('capture', 0))} 条，前者略高。"
        ),
        "",
        "三、月度走势",
        f"2026-01：{int(monthly_counts['2026-01'])} 篇，作为基准月。",
        f"2026-02：{int(monthly_counts['2026-02'])} 篇，较 1 月变动 {int(monthly_delta['2026-02']):+d} 篇。",
        f"2026-03：{int(monthly_counts['2026-03'])} 篇，较 2 月变动 {int(monthly_delta['2026-03']):+d} 篇。",
        "",
        "各月前三标签：",
    ]

    for month in ["2026-01", "2026-02", "2026-03"]:
        counts = top_labels_by_month.get(month, pd.Series(dtype="int64"))
        summary = "，".join(f"{label} {int(value)} 次" for label, value in counts.items())
        lines.append(f"{month}：{summary}")

    lines.extend(
        [
            "",
            "四、简要判断",
            "从季度节奏看，新闻总量没有剧烈波动，说明 CCUS 议题保持持续曝光，而非由单一事件驱动。",
            "从结构看，“碳移除”与“碳捕集”形成双核心，其中移除相关话题更稳定，捕集相关话题更容易受项目签约、融资或政策节点影响。",
            "“碳市场”“政策”“碳利用”等标签占比次一级，说明产业化与制度建设议题已形成辅助支撑，但尚未超过技术与项目本体的关注度。",
        ]
    )
    return lines


def style_sheet(ws) -> None:
    for row in ws.iter_rows():
        for cell in row:
            cell.font = BODY_FONT
            cell.alignment = Alignment(vertical="top", wrap_text=True)


def write_summary_sheet(ws, summary_lines: list[str]) -> None:
    ws.title = "摘要说明"
    ws.merge_cells("A1:D1")
    ws["A1"] = "CCUS 2026年一季度中文分析报告"
    ws["A1"].font = TITLE_FONT
    ws["A1"].fill = TITLE_FILL
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 24

    row = 3
    for line in summary_lines:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
        cell = ws.cell(row=row, column=1, value=line)
        if line and line[1:2] == "、":
            cell.font = SECTION_FONT
            cell.fill = SECTION_FILL
        else:
            cell.font = BODY_FONT
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        row += 1

    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 28
    ws.column_dimensions["C"].width = 28
    ws.column_dimensions["D"].width = 28


def write_tag_sheet(ws, tag_counts: pd.DataFrame) -> None:
    ws.title = "标签统计"
    headers = ["英文标签", "中文标签", "提及次数", "占比"]
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = SECTION_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center")

    for idx, row in enumerate(tag_counts.itertuples(index=False), start=2):
        ws.cell(row=idx, column=1, value=row.label)
        ws.cell(row=idx, column=2, value=row.label_zh)
        ws.cell(row=idx, column=3, value=int(row.news_count))
        pct_cell = ws.cell(row=idx, column=4, value=float(row.share_pct))
        pct_cell.number_format = "0.0%"

    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 12
    ws.freeze_panes = "A2"

    max_row = len(tag_counts) + 1
    category_ref = Reference(ws, min_col=2, min_row=2, max_row=max_row)
    data_ref = Reference(ws, min_col=3, min_row=1, max_row=max_row)

    bar_chart = BarChart()
    bar_chart.type = "bar"
    bar_chart.style = 10
    bar_chart.title = "CCUS 各领域新闻提及次数"
    bar_chart.y_axis.title = "领域"
    bar_chart.x_axis.title = "提及次数"
    bar_chart.height = 8
    bar_chart.width = 16
    bar_chart.add_data(data_ref, titles_from_data=True)
    bar_chart.set_categories(category_ref)
    bar_chart.legend = None
    bar_chart.dLbls = DataLabelList()
    bar_chart.dLbls.showVal = True
    ws.add_chart(bar_chart, "F2")

    pie_chart = PieChart()
    pie_chart.style = 10
    pie_chart.title = "CCUS 各领域占比"
    pie_chart.height = 10
    pie_chart.width = 12
    pie_chart.add_data(Reference(ws, min_col=3, min_row=1, max_row=max_row), titles_from_data=True)
    pie_chart.set_categories(category_ref)
    pie_chart.dLbls = DataLabelList()
    pie_chart.dLbls.showPercent = True
    pie_chart.dLbls.showLeaderLines = True
    pie_chart.dLbls.showLegendKey = False
    ws.add_chart(pie_chart, "F20")


def write_monthly_sheet(
    ws,
    monthly_counts: pd.Series,
    top_labels_by_month: dict[str, pd.Series],
    monthly_label_dist: pd.DataFrame,
) -> None:
    ws.title = "月度趋势"
    headers = ["月份", "新闻数量", "前三标签摘要"]
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = SECTION_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center")

    months = list(monthly_counts.index)
    for idx, month in enumerate(months, start=2):
        ws.cell(row=idx, column=1, value=month)
        ws.cell(row=idx, column=2, value=int(monthly_counts[month]))
        summary = "，".join(f"{label} {int(value)} 次" for label, value in top_labels_by_month.get(month, pd.Series(dtype='int64')).items())
        ws.cell(row=idx, column=3, value=summary)

    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 40
    ws.freeze_panes = "A2"

    line_chart = LineChart()
    line_chart.style = 10
    line_chart.title = "CCUS 月度新闻走势"
    line_chart.y_axis.title = "新闻数量"
    line_chart.x_axis.title = "月份"
    line_chart.height = 8
    line_chart.width = 16
    line_chart.add_data(Reference(ws, min_col=2, min_row=1, max_row=len(months) + 1), titles_from_data=True)
    line_chart.set_categories(Reference(ws, min_col=1, min_row=2, max_row=len(months) + 1))
    line_chart.dLbls = DataLabelList()
    line_chart.dLbls.showVal = True
    ws.add_chart(line_chart, "E2")

    start_row = 8
    start_col = 5
    ws.cell(row=start_row, column=start_col, value="标签月度分布数据").font = SECTION_FONT
    ws.cell(row=start_row, column=start_col).fill = SECTION_FILL

    ws.cell(row=start_row + 1, column=start_col, value="标签").font = SECTION_FONT
    ws.cell(row=start_row + 1, column=start_col).fill = HEADER_FILL
    for offset, month in enumerate(monthly_label_dist.columns, start=1):
        cell = ws.cell(row=start_row + 1, column=start_col + offset, value=month)
        cell.font = SECTION_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center")

    for row_offset, (label, values) in enumerate(monthly_label_dist.iterrows(), start=2):
        ws.cell(row=start_row + row_offset, column=start_col, value=label)
        for col_offset, value in enumerate(values.tolist(), start=1):
            ws.cell(row=start_row + row_offset, column=start_col + col_offset, value=int(value))

    dist_chart = LineChart()
    dist_chart.style = 10
    dist_chart.title = "各标签月度分布"
    dist_chart.y_axis.title = "提及次数"
    dist_chart.x_axis.title = "月份"
    dist_chart.height = 10
    dist_chart.width = 18
    data_min_row = start_row + 1
    data_max_row = start_row + 1 + len(monthly_label_dist)
    data_max_col = start_col + len(monthly_label_dist.columns)
    dist_chart.add_data(
        Reference(ws, min_col=start_col + 1, min_row=data_min_row, max_col=data_max_col, max_row=data_max_row),
        from_rows=True,
        titles_from_data=True,
    )
    dist_chart.set_categories(
        Reference(ws, min_col=start_col + 1, min_row=start_row + 1, max_col=data_max_col, max_row=start_row + 1)
    )
    dist_chart.legend.position = "r"
    ws.add_chart(dist_chart, "J2")


def write_news_sheet(ws, news: pd.DataFrame) -> None:
    ws.title = "新闻明细"
    output = news.copy()
    output["labels_zh"] = (
        output["labels"]
        .fillna("")
        .str.split("|")
        .apply(lambda items: "|".join(to_zh_label(item) for item in items if item))
    )
    output["source_categories_zh"] = (
        output["source_categories"]
        .fillna("")
        .str.split("|")
        .apply(lambda items: "|".join(SOURCE_CATEGORY_ZH.get(item, item) for item in items if item))
    )

    headers = ["日期", "标题", "原始标签", "中文标签", "链接", "原始来源类别", "中文来源类别"]
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = SECTION_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center")

    for row_idx, row in enumerate(output.itertuples(index=False), start=2):
        ws.cell(row=row_idx, column=1, value=row.date.strftime("%Y-%m-%d"))
        ws.cell(row=row_idx, column=2, value=row.title)
        ws.cell(row=row_idx, column=3, value=row.labels)
        ws.cell(row=row_idx, column=4, value=row.labels_zh)
        ws.cell(row=row_idx, column=5, value=row.url)
        ws.cell(row=row_idx, column=6, value=row.source_categories)
        ws.cell(row=row_idx, column=7, value=row.source_categories_zh)

    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 70
    ws.column_dimensions["C"].width = 24
    ws.column_dimensions["D"].width = 24
    ws.column_dimensions["E"].width = 80
    ws.column_dimensions["F"].width = 20
    ws.column_dimensions["G"].width = 20
    ws.freeze_panes = "A2"


def build_workbook(
    news: pd.DataFrame,
    tag_counts: pd.DataFrame,
    monthly_counts: pd.Series,
    top_labels_by_month: dict[str, pd.Series],
    source_counts: pd.Series,
    monthly_label_dist: pd.DataFrame,
) -> Workbook:
    wb = Workbook()
    summary_ws = wb.active

    summary_lines = build_summary_lines(news, tag_counts, monthly_counts, top_labels_by_month, source_counts)
    write_summary_sheet(summary_ws, summary_lines)
    write_tag_sheet(wb.create_sheet(), tag_counts)
    write_monthly_sheet(wb.create_sheet(), monthly_counts, top_labels_by_month, monthly_label_dist)
    write_news_sheet(wb.create_sheet(), news)

    for ws in wb.worksheets:
        style_sheet(ws)

    return wb


def main() -> None:
    news, tag_counts, monthly_counts, top_labels_by_month, source_counts, monthly_label_dist = prepare_data()
    workbook = build_workbook(
        news,
        tag_counts,
        monthly_counts,
        top_labels_by_month,
        source_counts,
        monthly_label_dist,
    )
    workbook.save(REPORT_XLSX)
    print(f"Generated: {REPORT_XLSX.name}")


if __name__ == "__main__":
    main()
