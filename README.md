# Excel 表格数据清洗与规则筛查工具（VBA）

一个零依赖的 VBA 模块，用来处理「一张主表带多行明细」的扁平表格：分组补齐明细键、生成派生列、按数值区间着色、按关键词与数值条件整组标记，最后对标记结果做一次反向校验。

所有列位置按**表头文本**查找，所有阈值、关键词、比较方向和配色都从可选的 `Settings` 工作表读取。代码里不含任何行业参数，换一张表只需要改配置。

## 能做什么

| 宏 | 作用 |
| --- | --- |
| `NormalizeDetailKeys` | 按组键分组；空白明细键数量等于期望行数时顺序编号，否则填入复查标记并着色 |
| `BuildDerivedColumns` | 在表格右侧追加三列：行金额、分组行数、分组金额（COUNTIF / SUMIF 公式） |
| `ApplyBandHighlight` | 把落在两个数值区间内的单元格分别着色 |
| `FlagGroupsByRule` | 某组内只要有一行同时满足「描述命中关键词」与「数值满足条件」，就把整组所有行标记 |
| `VerifyGroupFlags` | 重新读取被标记的行，报告不满足规则却仍被标记的误报数量 |
| `RunCleaningPipeline` | 依次执行以上四步 |

比较方向（`>=`、`>`、`<`、`<=`、`=`）是配置项，因此「达到阈值」和「低于阈值」两类筛查共用同一段代码，不需要各写一份。

## 快速开始

1. 在 Excel 或 WPS 表格中新建工作簿，按下表准备数据：第 1 行是表头，数据从第 2 行开始，**同一组的行需连续**。
2. `Alt + F11` 打开 VBA 编辑器 → `文件 → 导入文件` → 选择 `src/DataCleaner.bas`。
3. 可选：新建名为 `Settings` 的工作表，A 列写配置键、B 列写值，全部配置键见 [docs/Settings.md](docs/Settings.md)。不建这张表也能运行，会使用内置占位值。
4. 回到数据表，`Alt + F8` 运行 `RunCleaningPipeline`，或单独运行其中某一步。

### 最小表头

| RecordID | LineKey | ItemCount | Description | Quantity | UnitPrice |
| --- | --- | --- | --- | --- | --- |
| G-001 |  | 3 | SAMPLE_A part | 2 | 150 |
| G-001 |  | 3 | Generic item | 1 | 80 |
| G-001 |  | 3 | SAMPLE_B part | 4 | 60 |
| G-002 |  | 2 | Generic item | 1 | 300 |

`examples/sample.csv` 是可以直接导入的完整虚构数据，用来跑通全流程。

> 提示：`GroupByHeader` 决定 COUNTIF / SUMIF 以哪一列为条件。默认指向明细键列，明细键编号后每行唯一，因此结果等于该行自身值；把该配置改成组键列名即可得到真正的按组汇总。

## 设计说明

- **不写死列字母**：所有列通过表头文本定位，插入或删除列不会让代码失效。
- **配置驱动**：阈值、关键词、比较方向、颜色、提示文案全部外置；内置值只是占位示例，没有业务含义。
- **状态可恢复**：批量操作前关闭屏幕刷新与自动计算，出错时也会还原 `Application` 状态。
- **两阶段标记 + 反向校验**：先收集符合条件的组，再统一着色，最后复查误报，避免边读边写造成的漏标。
- **兼容性**：`Scripting.Dictionary` 采用后期绑定，无需勾选任何引用；适用于 Microsoft Excel 与 WPS 表格（Windows，需启用宏）。

## 目录结构

```
sheet-keyweave/
├── README.md              本说明
├── LICENSE                MIT 许可
├── PUSH_GUIDE.md          首次推送到 GitHub 的操作说明
├── .gitignore             忽略本地中间产物与临时文件
├── .gitattributes         换行与文本类型约定
├── src/
│   └── DataCleaner.bas    VBA 源码（导入即可用）
├── docs/
│   └── Settings.md        配置键完整说明
└── examples/
    └── sample.csv         虚构示例数据
```

## 使用边界

本项目只提供通用的表格处理方法。内置阈值、关键词与示例数据均为占位值，使用者应按自身场景配置，并自行确认数据来源、使用方式以及发布行为的合规性。示例数据为虚构内容，不对应任何真实记录。

## 许可

[MIT](LICENSE)
