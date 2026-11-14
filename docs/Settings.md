# Settings 配置说明

`Settings` 工作表是**可选的**。新建一张名为 `Settings` 的工作表，**A 列写配置键、B 列写值**即可覆盖默认值；某个键不存在或值为空时，代码自动使用下表默认值。

所有内置默认值都是中性占位值，用于让示例数据跑通，不代表任何业务口径。

## 列定位（值是表头文本）

| 键 | 默认值 | 说明 |
| --- | --- | --- |
| `GroupHeader` | `RecordID` | 分组键列，同组行按此列判断连续 |
| `DetailHeader` | `LineKey` | 明细键列，`NormalizeDetailKeys` 填充此列 |
| `ItemCountHeader` | `ItemCount` | 期望行数，取每组第一行 |
| `DescriptionHeader` | `Description` | 参与关键词匹配的文本列 |
| `QuantityHeader` | `Quantity` | 数量列 |
| `UnitPriceHeader` | `UnitPrice` | 单价列 |
| `GroupByHeader` | `LineKey` | COUNTIF / SUMIF 的条件列；改成组键列名可做按组汇总 |
| `MeasureHeader` | `GroupAmount` | 区间着色与规则判定读取的数值列 |
| `FlagTargetHeader` | `RecordID` | 被着色标记的列 |
| `LineAmountHeader` | `LineAmount` | 追加列：行金额的表头 |
| `GroupCountHeader` | `GroupRowCount` | 追加列：分组行数的表头 |
| `GroupSumHeader` | `GroupAmount` | 追加列：分组金额的表头 |

## 规则参数

| 键 | 默认值 | 说明 |
| --- | --- | --- |
| `Keywords` | `SAMPLE_A,SAMPLE_B` | 逗号分隔的关键词，大小写不敏感；留空则不会命中任何组 |
| `Comparison` | `>=` | 比较方向，可用 `>=` `>` `<` `<=` `=`（也接受 `GE` `GT` `LT` `BELOW` 等别名） |
| `LowerBound` | `1000` | 区间着色下限（含） |
| `UpperBound` | `2000` | 区间着色上限（含），超过上限另用一种颜色 |
| `RuleBound` | `1000` | 整组标记的数值阈值 |

## 文案与外观

| 键 | 默认值 | 说明 |
| --- | --- | --- |
| `ReviewNote` | `REVIEW_REQUIRED` | 明细键无法自动补齐时填入的提示 |
| `KeySeparator` | `-` | 明细键编号时的连接符 |
| `ColorReview` | `255,235,156` | 需复查单元格的底色 |
| `ColorDerived` | `226,235,248` | 追加派生列的底色 |
| `ColorBandLow` | `198,224,198` | 落在区间内的底色 |
| `ColorBandHigh` | `255,244,204` | 超过上限的底色 |
| `ColorFlag` | `144,200,160` | 整组标记的底色 |

颜色统一用 `R,G,B` 三个 0–255 的整数，逗号分隔。

## 配置示例

| A 列 | B 列 |
| --- | --- |
| `GroupHeader` | `RecordID` |
| `Keywords` | `SAMPLE_A,ALPHA` |
| `Comparison` | `<` |
| `RuleBound` | `800` |
| `GroupByHeader` | `RecordID` |
| `ColorFlag` | `144,200,160` |

把 `Comparison` 改成 `<`、`RuleBound` 改成 `800`，同一个 `FlagGroupsByRule` 就变成「低于 800 且命中关键词的组」筛查，无需改动代码。
