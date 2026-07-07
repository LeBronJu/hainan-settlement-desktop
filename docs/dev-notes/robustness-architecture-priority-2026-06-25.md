# 鲁棒性与解耦优先级（2026-06-25）

状态：历史架构审查和执行记录。部分建议已经完成；当前准绳请优先看 `docs/README.md`、`CONTEXT.md`、`docs/architecture.md` 和相关 current-behavior 文档。

## 背景

Win10/11 版、Win7/8 版、阶段一清洗、阶段二模板驱动生成、阶段二预检和校验报告已趋于稳定。下一轮工作应以数据正确性、可测试性和局部可维护性为主，不做大规模重写。

本轮体检不读取真实 Excel、真实台账、真实结算输出；验证优先使用纯 Core 测试、动态生成的极简 workbook、临时输出目录和脱敏样例。

## P0：结算正确性与输出安全

### 阶段二金额计算和差异判断

涉及文件：

- `src/HainanSettlementTool.Core/Models/DetailSettlementRow.cs`
- `src/HainanSettlementTool.Excel/Stage2SettlementGenerator.cs`

风险：

- 代理/居间明细的 `Gross`、`TaxAmount`、`CalculatedNet`、`ExpectedNet` 当前贴在 Excel 读取 Implementation 内。
- 台账金额和分表自算金额的差异判断也在 Excel 写表流程中产生。
- 后续调整税率、抵扣或公式缓存口径时，容易必须通过真实 workbook 才能验证。

建议：

- 先新增一个纯 Core Module，集中计算阶段二明细金额和差异判断。
- 用合成 `DetailSettlementRow` 或输入 DTO 写单元测试，不依赖 ClosedXML。
- Excel 层继续负责读取单元格和写 workbook，只调用 Core 规则得到金额和校验结果。

验收不变量：

- 现有阶段二输出 workbook 名称、报告名称和校验报告文本结构不变。
- 金额四舍五入仍使用当前 4 位小数口径。
- 不引入桌面 Excel 自动化。

### 文件访问和输出目录保护

涉及文件：

- `src/HainanSettlementTool.Core/Services/FileAccessGuard.cs`
- `src/HainanSettlementTool.Core/Services/HainanStage1Service.cs`
- `src/HainanSettlementTool.Core/Services/HainanStage2Service.cs`
- `src/HainanSettlementTool.Excel/*`

风险：

- 输出文件名和模板复制路径分散在阶段一、阶段二和 Excel 层。
- 后续若新增输出类型，可能绕过临时文件、占用文件或输出目录校验。

建议：

- 保持 `FileAccessGuard` 为统一入口，补齐最小测试。
- 后续新增输出文件前先检查是否需要 `RequireWritableWorkbook` 或目录创建规则。

## P1：大 Module 的 Locality 改善

### 阶段二生成器拆分

涉及文件：

- `src/HainanSettlementTool.Excel/Stage2SettlementGenerator.cs`

风险：

- 一个 Module 同时承担台账读取、预检、模板索引、分表写入、汇总表写入、报告输出。
- Public Interface 很小，但 Implementation 过集中；修改某个规则时需要理解整条阶段二链路。

建议拆分顺序：

1. 先抽纯 Core 金额规则和校验规则。
2. 再在 Excel 层内部拆出模板索引和预检读取 Module。
3. 再拆分表写入和汇总表写入 Module。
4. 最后才考虑报告输出 Module。

删除测试：

- 如果删除“预检 Module”，新增关系、新增客户、关键字段变化和模板读取失败的逻辑会重新散落回 generator。
- 如果删除“汇总表写入 Module”，插入月度区块、付款主体、抵扣和合计公式会重新和分表写入混在一起。

### 阶段一 raw detail 读取 Adapter

涉及文件：

- `src/HainanSettlementTool.Excel/RawDetailReader.cs`
- `src/HainanSettlementTool.Excel/CustomerCodeReader.cs`

风险：

- `.xlsx` / `.xls` / `.csv` 三种路径的行读取、列号、编码和空行处理分散在两个 reader。
- 新增格式或修复 `.xls` 兼容性时，容易只修电量清洗而漏掉户号读取。

建议：

- 抽一个内部 raw detail row Adapter，统一暴露客户名、户号、分时电量和 source row。
- `RawDetailReader` 和 `CustomerCodeReader` 只做投影，不各自理解文件格式。

验收不变量：

- 清洗后的电量处理表仍输出 `.xlsx`。
- `.csv` 默认编码兜底仍支持 GB18030。
- 未找到户号时仍由阶段一报告提示人工补齐。

### Win7/8 和 Win10/11 workflow 重复

涉及文件：

- `src/HainanSettlementTool.WinForms/MainForm.cs`
- `src/HainanSettlementTool.Wpf/MainWindow.xaml.cs`

风险：

- 两个 UI 都知道阶段一、只清洗、阶段二的执行顺序、确认点、日志摘要和 adapter 创建方式。
- 后续修阶段二预检或阶段一清洗流程，容易一边改了另一边漏掉。

建议：

- 暂不先动 UI 布局。
- 先抽共享 workflow Module，负责创建 `HainanStage1Service` / `HainanStage2Service`、执行阶段、返回摘要事件。
- UI 只负责确认、进度展示、日志展示和错误弹窗。

验收不变量：

- Win7/8 和 Win10/11 的按钮、文件选择和用户可见文案不因第一刀改变。
- 两个入口仍共享 Core/Excel 逻辑。

## P2：测试骨架和长期文档

### 最小测试项目

建议新增：

- `tests/HainanSettlementTool.Core.Tests/`
- 后续按需新增 `tests/HainanSettlementTool.Excel.Tests/`

优先测试：

- 阶段二金额计算和差异判断。
- `TextUtil.CustomerKey` / 数值解析等文本口径。
- `FileAccessGuard` 的临时文件、缺文件和输出目录行为。
- 阶段一 raw detail row Adapter 的合成输入，不读取真实 Excel。

暂缓：

- 大量真实 workbook fixture 自动化。
- 对模板格式做全量像素/样式级断言。

## 当前建议顺序

1. 创建最小 Core 测试项目。
2. 抽阶段二金额计算和差异判断到 Core，并测试。
3. 回到 `Stage2SettlementGenerator`，只替换金额计算调用，不改变 workbook 写入。
4. 抽 raw detail row Adapter，统一 `.xlsx/.xls/.csv` 读取投影。
5. 视稳定性再拆阶段二预检/模板索引/汇总写入。

## 不做的事

- 不读取真实 Excel、真实台账或真实结算输出作为自动化测试输入。
- 不为了“架构好看”大拆 UI。
- 不在没有 fixture 保护时重写阶段二 workbook 模板生成。
- 不引入新的桌面框架或 Excel 自动化依赖。

## 本轮执行记录

- 已新增 `tests/HainanSettlementTool.Core.Tests/` 最小测试项目。
- 已新增 `Stage2SettlementCalculator`，集中阶段二明细金额计算、金额格式化、台账与分表自算差异判断和差异问题生成。
- `Stage2SettlementGenerator` 已改为调用 Core 规则，仍保留 Excel 读取、模板复制、公式写入和报告输出 Implementation。
- 已为 `FileAccessGuard` 补齐输出目录创建、Excel 临时文件拦截、输出文件占用提示的最小测试；生产实现无需调整。
- 已新增 `tests/HainanSettlementTool.Excel.Tests/`，用合成 `.csv` / `.xlsx` 验证阶段一原始零售侧明细读取。
- 已新增内部 `RawDetailRowReader` Adapter，统一 `.xlsx` / `.xls` / `.csv` 的行读取、列解析和 CSV 编码兜底；`RawDetailReader` 与 `CustomerCodeReader` 只保留投影逻辑。
- `CustomerCodeReader` 现在也能从 `.csv` 原始零售侧明细读取户号，避免“只清洗电量能读 CSV、补户号不能读 CSV”的漂移。
- 审查后补强：同一客户若在原始明细里出现多个不同户号，不自动选择第一条，而是不给出户号映射，让阶段一报告继续提示人工补齐；同一客户重复且户号一致仍可自动补齐。
- 已为 `LedgerStage1Updater` 补齐合成 workbook 测试，覆盖新增客户插入 footer 前、唯一户号补齐/冲突户号不自动补、输出台账副本不覆盖基础台账。
- 已新增 Core `SettlementWorkflow`，集中阶段一、只清洗和阶段二的执行结果摘要；Win7/8 与 Win10/11 UI 继续保留确认、进度、弹窗和错误展示，但不再各自拼接相同的成功摘要。
- 已深化 Core `SettlementWorkflow` 的阶段二流程：新增阶段二预检计划和完成结果，WPF 只负责展示预检确认，Core 负责确认后继续/取消的生成决策，并用 Core 测试覆盖取消时不生成。
