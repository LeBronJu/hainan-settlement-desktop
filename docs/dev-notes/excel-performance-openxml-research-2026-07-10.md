# Excel 批处理性能与 Open XML SDK 迁移研究

状态：当前架构 note

更新日期：2026-07-10

## 目的

记录广东批量分表实测后的性能研究、海南优先的优化顺序，以及当前项目引入 Open XML SDK 的收益、风险和渐进路线。该文档是当前技术研究准绳，不改变已验证的结算规则。

## 触发背景

- 用户已对广东分表月份初始化做简单验收，真实批次共 600 多个 workbook，总耗时 4 分多钟。
- 该性能已满足当天广东临时需求；广东完整结算功能和性能扩展暂不进入下一个实现切片。
- 用户决定性能研究先从自己负责、规则熟悉、数量较少的海南流程开始，稳定后再评估重庆。
- 海南适合做正确性和迁移方法试点，但其数据量不能单独证明广东大批量吞吐量；后续需用合成 workbook 复制批次做压测。

## 当前技术事实

- Excel 项目目标框架是 `.NET Framework 4.7.2`。
- 顶层 Excel 依赖是 `ClosedXML 0.104.2` 和 `ExcelDataReader 3.7.0`。
- ClosedXML 已传递引入 `DocumentFormat.OpenXml 3.1.1`；试验 Open XML SDK 时可先添加显式直接引用，不需同时更换 UI 或 Core 合同。
- WPF 使用 `Task.Run` 把 Excel 工作放到后台线程，但当前 workbook 循环仍是串行；“界面不卡死”不等于“多 workbook 并行”。
- 广东当前路径会在 WPF 预检后由 Generate 再做一次完整 Analyze；需修改的 workbook 还会先完整复制，再重新打开输出副本，保存时请求公式评估。这些是已确认的重复工作，但广东优化当前暂缓。
- 海南阶段二 Analyze 和 Generate 也会分别读取台账、分表模板和汇总表。进入性能优化前应先测量，不凭感觉替换 Excel 引擎。

## 当前研究结论

1. 不做 ClosedXML 到 Open XML SDK 的一次性全量重写。
2. 候选终态是混合 Excel Adapter：Open XML SDK 用于大批量扫描、只读投影和简单定点修改；ClosedXML 暂时保留复杂模板复制、插行/插列、sheet 关系和公式缓存写入。
3. 先做不更换引擎的基线优化：计时、消除重复读取、减少不必要保存，再判断 Open XML SDK 的真实边际收益。
4. 一次运行中，同一 workbook 的业务修改应集中完成，原则上只做一次生成阶段打开和一次保存，不得随功能增加反复重开 workbook。
5. 并行粒度应是“独立 workbook”，并使用受控并发；不共享 `XLWorkbook` / `SpreadsheetDocument` 实例，不一次将大量 workbook 留在内存。

这是当前推荐方向，不是“已决定最终移除 ClosedXML”的 ADR。只有当两种 Adapter 都经过真实验证且迁移收益足够时，才决定是否继续替换。

## Open XML SDK 在本项目的具体风险

### 1. 公式计算与缓存值

Open XML SDK 可以写公式文本和上次计算缓存，但不是 Excel 公式计算引擎。当前海南阶段二和员工奖励保存时会请求 ClosedXML 评估公式。如果改为只写公式而不更新缓存，WPS/Excel 重算前可能显示旧金额，下游程序若只读缓存也可能读到旧值。

### 2. 插行、插列和公式引用迁移

海南阶段一会插入新客户行和复制月份块；阶段二汇总表会插入六列月份块。Open XML SDK 不会自动为业务代码修正所有公式引用、合并区域、条件格式、数据验证、命名范围、打印区域和图形锚点。结构性写入是最高风险区。

### 3. sheet 克隆和关系图

阶段二分表会复制、删除和重命名月份 sheet。worksheet XML 之外还可能存在批注、图片、超链接、表格、打印设置、关系 ID 和其它 part。只复制 `WorksheetPart` 或 XML 内容不等于完整克隆 sheet。

### 4. 样式、隐藏状态和显示值

本项目依赖隐藏列、列宽、边框、日期格式、合并表头和空白样式。Open XML 样式通过 workbook 级索引引用，日期常是“数字 + 样式”，文本可能来自 shared strings。ClosedXML 的 `GetFormattedString()` 等便利行为需要在 Open XML Adapter 中显式实现和测试。

### 5. `.xls` / `.csv` 输入不能全量迁移

海南原始零售侧明细支持 `.xlsx/.xls/.csv`。Open XML SDK 只处理 Office Open XML 容器，不会替代旧 `.xls` 和 CSV 路径。即使长期大量使用 Open XML SDK，`ExcelDataReader` 和 CSV Adapter 仍需保留。

### 6. Excel / WPS 兼容和未知扩展 part

生产 workbook 可能带 Office/WPS 扩展标记。定点修改应尽量保留未触及的 part；重建 workbook 或不完整克隆可能导致 Excel/WPS 提示“已修复文件”，或造成两者显示差异。

### 7. 性能收益不是自动获得

Open XML SDK 的 DOM 模式仍会加载完整 XML part。对大表的优势通常来自 `OpenXmlReader` 的前向只读、定点 XML 修改和避免公式全量评估，而不是仅仅替换 NuGet 包。

## 海南试点顺序

### 第一批：基线和无行为优化

1. 分别记录海南阶段一、阶段二、员工奖励的扫描、解析、计算、写入和保存耗时。
2. 记录每个 workbook 的耗时和文件大小，但不在 WPF 日志中密集输出客户级信息。
3. 复用预检产生的轻量处理计划；用文件长度和 `LastWriteTimeUtc` 检测用户在确认窗口期间是否修改源文件，只重新分析已变化文件。
4. 公共台账和其它共享输入每轮只读取一次；所有业务修改集中后只保存一次。
5. 优先证明正确性和实际瓶颈，不在没有数据时先调高并发度。

### 第二批：Open XML 只读影子 Adapter

优先候选：

- `HainanStage2LedgerReader`：只读台账并投影为阶段二明细。
- `HainanStage2TemplateIndex`：扫描分表 sheet 和模板标识。
- 海南员工奖励的台账读取部分。
- `HainanRawDetailRowReader` 的 `.xlsx` 分支；`.xls/.csv` 继续使用现有 Adapter。

影子期由 ClosedXML 仍产生正式结果，Open XML Adapter 同时读取并比较归一化 DTO。只有差异为零并通过合成/授权回放后，才切换正式读取。

### 最后评估的高风险 Writer

- `HainanStage1LedgerUpdater`：插入客户行、复制月份块、复制样式和隐藏列。
- `HainanStage2SplitWorkbookWriter`：复制/删除 sheet、伸缩明细行、写入公式和修复样式。
- `HainanStage2SummaryWorkbookWriter`：插入月份列块、隐藏历史列、重写公式和合并区域，是最高风险模块。
- `HainanEmployeePowerRewardGenerator` 写出部分：从零构建样式、合并区域和公式，输出数量少，当前迁移收益低。

## 验收门槛

任何正式切换必须同时满足：

- 合成 workbook 测试覆盖隐藏列/sheet、合并区域、公式与缓存、日期格式、样式索引和空白单元格。
- ClosedXML 和 Open XML 读取结果按业务 DTO 对比无差异。
- 生成 workbook 通过 Open XML 结构验证，Excel 和 WPS 打开无修复提示。
- 金额、公式文本、缓存值、隐藏状态、列宽、合并表头、签字/页脚和 sheet 顺序符合现有验收不变量。
- 继续保持源 workbook 只读，输出副本与单文件失败隔离。
- 同一批输入上记录优化前后耗时、峰值内存和失败数；没有显著收益时不为“已使用新 SDK”承担迁移风险。

## 官方技术参考

- Open XML SDK 概览：`https://learn.microsoft.com/en-us/office/open-xml/open-xml-sdk`
- 大型 Spreadsheet 的 DOM/SAX 读取差异：`https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/how-to-parse-and-read-a-large-spreadsheet`
- SpreadsheetML 公式与缓存值：`https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-formulas`
- Calculation Chain 的作用和限制：`https://learn.microsoft.com/en-us/office/open-xml/spreadsheet/working-with-the-calculation-chain`

## 下一个实现切片

1. 只改海南 WPF/Core/Excel 共享主线，不为冻结的 WinForms 增加 UI 工作。
2. 先添加计时与基线报告，不改 workbook 输出。
3. 测量后优先消除重复预检/读取和不必要公式评估。
4. 再实现第一个 Open XML 只读影子 Adapter，不先动阶段二汇总 Writer。
5. 未获得新的明确真实数据授权前，只用合成 workbook 和已有脱敏 fixture；历史授权不自动延续。
