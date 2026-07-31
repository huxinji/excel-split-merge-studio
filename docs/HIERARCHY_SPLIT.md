# 层级筛选拆分 / Hierarchy Multi-Select Split

## 中文

层级筛选拆分用于存在明确上下级关系的字段，例如：

```text
大区 → 片区 → 网点
```

最多配置 3 级。最深一级是本次拆分目标，前面的每一级用于限定数据范围，并且都支持
多选。范围值从完整 Sheet 后台读取，而不是只读取预览中的前几行。

### 多选关系

- 同一级内是“或”：大区选择“华东、华南”表示保留任一大区；
- 不同级之间是“且”：大区多选与片区多选必须同时满足；
- 候选值级联：先选择多个大区后，片区选择器展示这些大区下片区的并集；
- 同名子级隔离：不同大区下即使存在同名片区，也会按完整路径生成独立输出。

### 示例一：多选大区，按片区拆分

1. 在拆分工作簿中先选择数据源。
2. 将拆分方式切换为“层级筛选拆分”。
3. 选择“2 级：多选上级，按第 2 级拆分”。
4. 第 1 级字段选择“大区”，点击“多选”，勾选“华东、华南”。
5. 第 2 级字段选择“片区”。
6. 设置输出目录并开始拆分。

程序保留两个大区的记录，再按完整的“大区 → 片区”路径输出。如果两个大区都存在
“核心片区”，会分别输出“华东__核心片区”和“华南__核心片区”。

### 示例二：多选大区和片区，按网点拆分

1. 选择“3 级：多选前两级，按第 3 级拆分”。
2. 第 1 级字段选择“大区”，勾选“华东、华南”。
3. 第 2 级字段选择“片区”，勾选“上海、广东”。
4. 第 3 级字段选择“网点”。

筛选关系为：

```text
大区 ∈ {华东, 华南} AND 片区 ∈ {上海, 广东}
```

匹配记录再按完整的“大区 → 片区 → 网点”路径拆分。

### 多选器操作

- 输入关键词实时搜索；
- “全选当前结果”只勾选当前搜索结果；
- “清空选择”取消全部勾选；
- 列表最多展示前 500 个不同值，仍可手工添加列表外的精确值；
- 已选值在主界面显示摘要，完整列表可通过悬停查看。

### 数据与输出规则

- 各层字段不能重复；
- 范围值采用去除首尾空格、统一全半角和忽略大小写后的精确匹配；
- 空值可单独分组或跳过；
- 默认不覆盖源文件或已有输出；
- 核对关系为：输入行数 = 输出行数 + 层级筛选排除行数；
- 旧的单值层级配置仍然兼容；
- 命名模板可以使用 `{hierarchy_path}`、`{split_field}`、`{split_value}`、
  `{level1_value}`、`{level2_value}` 和 `{level3_value}`。

## English

Hierarchy Multi-Select Split supports fields with a clear parent-child structure, for example:

```text
Area → District → Branch
```

Up to three levels can be configured. The deepest level is the split target. Every preceding
level limits the data scope and supports multiple selections. Scope values are discovered from
the complete worksheet in a background thread.

### Selection semantics

- Values within one level are OR-ed.
- Different hierarchy levels are AND-ed.
- Candidate values cascade: selecting multiple Areas exposes the union of their Districts.
- Same-named children under different parents are grouped by their full path and remain separate.

### Example 1: select multiple areas and split by district

1. Select the data source first in the Split Workbooks workspace.
2. Switch the split method to **Hierarchy**.
3. Choose **2 levels: select parent values, split by level 2**.
4. Set level 1 to `Area`, open **Select…**, and check `East` and `South`.
5. Set level 2 to `District`.
6. Choose an output folder and start.

If both Areas contain a `Core` District, the application creates separate `East__Core` and
`South__Core` outputs.

### Example 2: multi-select two parent levels, then split by branch

1. Choose **3 levels: select two parent scopes, split by level 3**.
2. Select `East` and `South` at level 1.
3. Select `Shanghai` and `Guangdong` at level 2.
4. Set level 3 to `Branch`.

The rule is:

```text
Area IN {East, South} AND District IN {Shanghai, Guangdong}
```

Matching rows are split by their complete Area → District → Branch path.

### Selector controls

- Search values as you type.
- **Select Visible** checks only the current search results.
- **Clear Selection** removes all checks.
- Up to 500 values are listed; any other exact value can be added manually.
- The main workspace shows a compact selection summary and a full tooltip.

### Data and output rules

- A field cannot be reused at multiple levels.
- Exact matching normalizes whitespace, character width, and letter case.
- Blank target values can be grouped or skipped.
- Source files and existing outputs are not overwritten by default.
- Reconciliation uses: input rows = output rows + hierarchy-filtered rows.
- Legacy single-value hierarchy configurations remain compatible.
- Naming templates support `{hierarchy_path}`, `{split_field}`, `{split_value}`,
  `{level1_value}`, `{level2_value}`, and `{level3_value}`.
