# CO2 11类重分类说明

## 本次做了什么

基于上次生成的 `filtered_news/all_literature_merged_dedup.csv`，新增了脚本 `reclassify_co2_news.py`，不再重新合并 Excel，直接对已有 CSV 进行 11 类重分类。

目标分类为：

- CO₂光转化
- CO₂热转化
- CO₂电转化
- CO₂矿化
- CO₂生物转化
- CO₂捕集与分离
- CO₂运输与封存
- 系统集成与工艺耦合
- 能源经济与LCA
- 政策与产业化
- 其他

## 脚本调整记录

1. 新增 `reclassify_co2_news.py`
   - 默认输入：`filtered_news/all_literature_merged_dedup.csv`
   - 默认输出：
     - `filtered_news/all_literature_reclassified_11cats.csv`
     - `filtered_news/all_literature_reclassified_11cats.xlsx`

2. 分类方法
   - 使用 `sentence-transformers/all-MiniLM-L6-v2`
   - 先做向量相似度分类
   - 再叠加关键词规则做校正

3. 阈值调整
   - 初版阈值较高，导致很多边界样本落入“其他”
   - 后续将默认阈值下调到 `0.42`

4. 关键词规则增强
   - `MOF/COF` 不再一律归入“CO₂捕集与分离”
   - 现在优先判断其用途：
     - `MOF/COF + photocatalysis` -> `CO₂光转化`
     - `MOF/COF + electroreduction/electrocatalysis` -> `CO₂电转化`
     - `MOF/COF + thermocatalysis` -> `CO₂热转化`
     - `MOF/COF + capture/separation/adsorption/membrane` -> `CO₂捕集与分离`

5. 边界样本尽量回收进11类
   - 对有明确光/热/电/生物转化语义的有机合成/利用类文章，优先归入相应“利用”类别
   - 增加了以下规则类目：
     - 运输与封存关键词
     - LCA/TEA关键词
     - 政策/产业化关键词
     - 系统集成/工艺耦合关键词
     - 矿化关键词

6. 误判修正
   - 修复了 `terminal` 命中 `terminally` 导致误判为“CO₂运输与封存”的问题
   - 英文关键词匹配改为按词边界匹配，而不是简单子串匹配
   - 对“明确是转化利用，但摘要里顺带出现 capture”的文章，转化路径优先于捕集标签

## 当前输出结果

当前重分类结果文件：

- `filtered_news/all_literature_reclassified_11cats.csv`
- `filtered_news/all_literature_reclassified_11cats.xlsx`

当前一版的分类数量为：

- CO₂光转化：42
- CO₂热转化：13
- CO₂电转化：71
- CO₂矿化：3
- CO₂生物转化：4
- CO₂捕集与分离：39
- CO₂运输与封存：8
- 系统集成与工艺耦合：6
- 能源经济与LCA：8
- 政策与产业化：20
- 其他：87

## 结果文件中的辅助字段

输出 CSV/XLSX 中增加了以下字段，便于回查：

- `assigned_category`：最终分类
- `assigned_score`：最终分类分数
- `rule_category`：若命中规则，显示规则指定的分类
- `rule_reason`：命中的规则原因

## 备注

当前版本已经明显减少了“其他”和若干边界误判，但仍然可以继续微调，例如：

- 提高 `CO₂矿化` 的召回
- 进一步区分“捕集材料研究”与“捕集后转化材料研究”
- 对多标签文章增加主标签/次标签机制
