# Excel Split & Merge Studio

面向 Excel 日常办公与批量数据处理的中英双语桌面软件。它将字段值拆分、最多三级的
管理层级多选拆分、纵向/横向合并、字段匹配关联、同名 Sheet 汇总和分组聚合集中在一个
响应式界面中。

A bilingual desktop application for everyday Excel operations: field-value splitting,
multi-select hierarchy splitting with up to three levels, vertical and horizontal merging,
key-based joins, same-name worksheet consolidation, and grouped aggregation.

![默认按字段值拆分界面](docs/screenshots/split-field-workspace.png)

![层级多选拆分界面](docs/screenshots/hierarchy-split-workspace.png)

## 核心功能 / Core Features

### 拆分 / Split

- “按指定字段值拆分”位于首屏并作为默认模式；
- “层级筛选拆分”支持 2～3 级管理层级；
- 每个上级范围均可搜索、多选、全选当前结果、清空或手工添加精确值；
- 可同时选择多个大区，再按各自片区拆分；
- 3 级模式可多选大区和片区，再按网点拆分；
- 多个父级下的同名子级按完整层级路径分别输出，不会误合并；
- 层级范围值从完整 Sheet 后台读取，不只依赖预览行；
- 每个不同字段值生成一个文件或一个 Sheet；
- 空字段值可单独分组，也可跳过；
- 按固定行数、指定份数或每个 Sheet 一个文件拆分；
- 可选择全部、可见、当前单个或自定义多选 Sheet；
- 支持 `.xlsx`、`.xlsm`、`.xls`、`.xlsb`、`.csv`、`.tsv`（旧格式需要可选依赖）；
- 输出自动编号，默认不覆盖已有文件。

- “Split by selected field value” remains the default above-the-fold workflow;
- hierarchy splitting supports two or three management levels;
- every parent scope supports searchable multi-selection, Select Visible, Clear, and exact custom values;
- select multiple Areas and split by their Districts, or select multiple Areas and Districts before splitting by Branch;
- same-named children under different parents remain separate path-based outputs;
- hierarchy values are read from the complete worksheet in a background thread;
- one output file or worksheet per distinct target path;
- blank values can be grouped or skipped;
- fixed-row, fixed-part, and one-file-per-sheet splitting;
- all, visible, current, or custom worksheet scope;
- safe auto-numbering instead of overwriting existing output by default.

See [层级筛选拆分说明 / Hierarchy Split Guide](docs/HIERARCHY_SPLIT.md).

### 合并 / Merge

- 纵向追加：严格一致、字段对齐、并集、交集、主表字段或按列位置；
- 关键字段 Join：Left、Right、Inner、Full Outer；
- 重复键策略、键值规范化、单字段模糊匹配和字段冲突处理；
- 多工作簿 Sheet 原样汇总；
- 同名 Sheet 跨文件汇总；
- 横向按行拼接，可设置来源前缀与空白列间距；
- 分组求和、计数、平均、最小、最大和去重计数；
- 全行或指定字段去重、来源追踪、任务报告与行数核对。

- vertical append with strict, aligned, union, intersection, master, or positional schemas;
- Left, Right, Inner, and Full Outer key joins;
- duplicate-key policies, normalized keys, fuzzy matching, and conflict handling;
- workbook-level and same-name-sheet consolidation;
- side-by-side concatenation with source prefixes and configurable gaps;
- grouped aggregation, deduplication, source tracking, task reports, and reconciliation.

## 交互 / UX

- 顶部“拆分工作簿 / 合并工作簿”切换；
- 中英语言实时切换；
- 左侧步骤卡片、右侧数据预览与日志、固定底部操作栏；
- 初始窗口按屏幕可用区域约 86% × 84% 居中打开，不默认最大化；
- 左侧内容可滚动，适配 1366×768、1920×1080 与系统 DPI 缩放；
- 文件扫描、范围值读取、预览和实际处理均在后台线程执行；
- 支持暂停、继续、取消、进度显示和打开输出目录。

- live Split/Merge workspace and Chinese/English switching;
- step cards on the left, preview and logs on the right, and a fixed action footer;
- centered responsive startup geometry without forced maximization;
- scrollable content for smaller displays and system DPI scaling;
- background scanning, hierarchy-value discovery, preview, and workbook processing;
- pause, resume, cancel, progress, and output-folder access.

## 项目结构 / Project Layout

```text
excel-split-merge-studio/
├─ src/excel_studio/
│  ├─ config/       # 中英文本、设置与常量
│  ├─ models/       # 拆分、合并、执行与结果模型
│  ├─ services/     # 扫描、读取、拆分、合并、输出与核对
│  ├─ ui/           # 唯一正式 PySide6 界面
│  └─ workers/      # 后台任务线程
├─ tests/           # 单元、集成、UI 与性能测试
├─ examples/        # 示例工作簿与规则模板
├─ installer/       # Windows 安装脚本
├─ scripts/         # 单文件构建、截图与打包验证
└─ .github/workflows/
```

## 数据安全与限制 / Data Safety & Limitations

- 源文件仅作为输入读取；默认策略不会覆盖源文件或已有输出。
- 层级范围值采用规范化后的精确匹配，不进行模糊归类。
- 多选父级时使用完整层级路径命名并分组，防止同名子级跨父级合并。
- 复杂图表、图片、外部链接、条件格式和 VBA 跨工作簿复制采用尽力保留。
- 公式不会由 Python 像桌面 Excel 一样完整重算。

- source workbooks are read-only inputs and existing outputs are auto-numbered by default;
- hierarchy scope values use normalized exact matching, not fuzzy classification;
- full path grouping keeps same-named children under different selected parents separate;
- complex workbook objects are preserved on a best-effort basis;
- Python does not fully recalculate formulas like desktop Excel;
