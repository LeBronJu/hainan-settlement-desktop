# 多省份代码质量主线（2026-07-08）

状态：当前任务 note。代码质量、海南命名中性化、大类拆分和多省份可维护性工作先读本文件；稳定架构规则仍以 `docs/architecture.md` 和 `docs/dev-notes/multi-province-readiness-2026-07-07.md` 为准。

## 当前主线

今日并行两条主线：

1. 重庆阶段二实机验收支援。用户测试 `HainanSettlementTool-Win10-11-Release-20260708-101959.zip` 时，如果反馈 bug、输出差异或需要重做的规则，优先级高于代码质量拆分。
2. 多省份代码质量提升。目标是控制大类、逐步修正误导性的海南/通用命名，并让后续 agent 能从文档知道下一步做什么。

本主线不读取真实 Excel、真实台账、客户数据、截图或结算结果，除非用户针对某个文件、目录或回放明确授权。当前拆分优先使用合成测试和现有单元测试。

## 固定工作协议

用户可能用长段自然语言反馈。Agent 应先把它解析成以下字段，再行动：

- `优先级`：实机测试 bug / 结算正确性疑问最高；代码质量和命名治理次之；纯体验优化再次。
- `授权级别`：`先看看`、`考虑`、`候选建议`、`工作计划` 表示只读分析，不改文件；`做吧`、`开始吧`、`授权你`、`我去休息/下班` 表示可按当前文档规则自主执行、验证、收尾。
- `数据边界`：未明确授权的真实业务文件不读；历史授权只作为背景，不自动变成本轮读权限。
- `文档边界`：长计划和执行连续性写入当前 dev note；`HANDOFF.md` 只保留当前状态、最新验证、包路径和下一步。
- `收尾格式`：每轮最终回复必须给出文档影响判断、验证、未做。

## 候选清单

2026-07-08 统计的主要大文件如下。WinForms 已冻结，除非用户明确重新开启 Win7/8 支持，不作为代码质量主线目标。

| 行数 | 文件 | 当前判断 |
| ---: | --- | --- |
| 1821 | `src/HainanSettlementTool.Excel/Stage2SettlementGenerator.cs` | 已处理第一批。重命名并拆为 `HainanStage2SettlementGenerator`、台账读取、模板索引、分表写入、汇总表写入、报告输出和工具组件。 |
| 1519 | `src/HainanSettlementTool.Wpf/MainWindow.xaml.cs` | 第二批已部分处理。已拆日志、进度、结果、弹窗、路径选择、输入状态和 options 构造；后续继续拆 workflow 编排和省份 UI 状态应用。 |
| 1207 | `src/HainanSettlementTool.WinForms/MainForm.cs` | 冻结。只因共享层修复被动受益，不安排主动拆分。 |
| 874 | `src/HainanSettlementTool.Excel/ChongqingPowerCleanGenerator.cs` | P2。重庆阶段一清洗稳定后可拆读取、校验、汇总、报告写入。 |
| 804 | `tests/HainanSettlementTool.Excel.Tests/Stage2SettlementGeneratorTests.cs` | 已随第一批重命名为 `HainanStage2SettlementGeneratorTests.cs`；后续再按行为分组拆测试。 |
| 803 | `src/HainanSettlementTool.Excel/ChongqingLedgerStage1Updater.cs` | P2。可按预检、客户决定应用、月份块写入、报告输出拆分。 |
| 721 | `tests/HainanSettlementTool.Excel.Tests/ChongqingPowerCleanGeneratorTests.cs` | 跟随重庆清洗拆分后再整理。 |
| 682 | `src/HainanSettlementTool.Excel/EmployeeRewardGenerator.cs` | P2。可按台账读取、汇总计算、workbook 输出拆分。 |
| 568 | `src/HainanSettlementTool.Wpf/MainWindow.xaml` | P2。先不为拆而拆，等 workflow/input 控制器稳定后再看 XAML 资源和控件分组。 |
| 544 | `src/HainanSettlementTool.Excel/ChongqingStage2SummaryWorkbookWriter.cs` | 观察。重庆阶段二刚完成首版，先等实机验收反馈，再拆长期字段、月度块、支付方 sheet。 |

## 命名治理规则

- `Hainan...` / `Chongqing...` 用于省份专属业务规则和 Excel 结构。
- `Province...` / `Settlement...` 只用于真实跨省共享模型、workflow、能力 profile 或 UI 壳。
- 不做项目名、解决方案名、程序集名、根命名空间和发布包名的一次性大迁移。
- 优先改最误导维护者的内部类名、变量名和 helper 名，例如实际只处理海南阶段二的通用名。
- 每个命名切片必须配套 focused 测试或构建验证，避免 XAML、脚本路径、反射字符串或项目引用被纯重命名破坏。

## 第一批目标

第一批只处理海南阶段二 Excel 生成器，不改业务口径：

- 将 `Stage2SettlementGenerator` 的海南属性显式化，避免与重庆阶段二或未来省份混淆。
- 按职责拆出内部组件，候选方向包括台账读取/分组、分表模板生成、汇总表写入、预检报告、Excel 工具函数。
- 保持 `ClosedXmlSettlementExcelGateway` 对 Core 的外部行为不变。
- 保持输出文件名、sheet 名、校验报告内容、支付方预检语义和异常提示不变。
- 同步整理对应测试文件，但不为了美观重写所有断言。

验收不变量：

- 不读取真实生产 workbook。
- 海南阶段二现有合成测试继续通过。
- Core/Excel 测试继续通过。
- WPF Debug/Release 编译通过。
- 文档 guardrail 和 `git diff --check` 通过。

## 第一批进展

2026-07-08 已完成海南阶段二 Excel 生成器第一轮拆分：

- `Stage2SettlementGenerator` 重命名为 `HainanStage2SettlementGenerator`，保留入口编排和预检构建。
- 新增 `HainanStage2LedgerReader`，负责读取海南阶段二台账代理/居间明细。
- 新增 `HainanStage2TemplateIndex`，负责扫描上月代理/居间分表模板。
- 新增 `HainanStage2SplitWorkbookWriter`，负责复制模板、准备月份 sheet、写分表明细、修合计行样式和签字日期。
- 新增 `HainanStage2SummaryWorkbookWriter`，负责汇总表副本、支付方继承/决策、新增汇总主体和汇总公式。
- 新增 `HainanStage2ReportWriter`，负责 JSON 总报告、自动生成提示和阶段二校验报告。
- 新增 `HainanStage2ExcelUtil` / `HainanStage2Models`，收纳共享常量、名称 key、支付方 override、通用 worksheet 工具和内部模型。

本轮只做结构拆分和海南命名显式化，不改输出文件名、公式、支付方规则或预检文案。

本轮验证：

- `dotnet test tests\HainanSettlementTool.Core.Tests\HainanSettlementTool.Core.Tests.csproj --no-restore` 通过，26 个测试。
- `dotnet test tests\HainanSettlementTool.Excel.Tests\HainanSettlementTool.Excel.Tests.csproj --no-restore` 通过，27 个测试。
- `dotnet msbuild src\HainanSettlementTool.Wpf\HainanSettlementTool.Wpf.csproj /restore /p:Configuration=Debug /m` 通过。
- `dotnet msbuild src\HainanSettlementTool.Wpf\HainanSettlementTool.Wpf.csproj /restore /p:Configuration=Release /m` 通过。

## 第二批进展

2026-07-08 已完成 WPF 主窗口输入状态第一轮拆分：

- 新增 `MainWindowInputController`，负责载入/保存用户输入路径、恢复已保存省份、构造海南阶段一/阶段二、重庆阶段一/阶段二和员工电量奖励 options、读取结算月份/奖励月份/当前省份，以及清空阶段输入。
- `MainWindow.xaml.cs` 不再直接拼装阶段 options 或直接序列化 `UserInputSnapshot`；窗口仍保留按钮事件、workflow 编排、进度/结果/弹窗协调和省份 UI 状态应用。
- `MainWindow.xaml.cs` 从 1519 行降到 1324 行。本轮不改变 UI 文案、输入字段、持久化文件格式或运行流程；同时补上原本“保存省份但不恢复省份”的输入状态缺口，并保护 `InitializeComponent()` 期间 SelectionChanged 早触发时的空引用风险。

本轮验证：

- `dotnet msbuild src\HainanSettlementTool.Wpf\HainanSettlementTool.Wpf.csproj /restore /p:Configuration=Debug /m` 通过。
- `dotnet test tests\HainanSettlementTool.Core.Tests\HainanSettlementTool.Core.Tests.csproj --no-restore` 通过，26 个测试。
- `dotnet test tests\HainanSettlementTool.Excel.Tests\HainanSettlementTool.Excel.Tests.csproj --no-restore` 通过，27 个测试。
- `dotnet msbuild src\HainanSettlementTool.Wpf\HainanSettlementTool.Wpf.csproj /restore /p:Configuration=Release /m` 通过。

## 待办

- [x] 建立本 current task note，并接入 `docs/README.md`。
- [x] 拆分并显式命名海南阶段二 Excel 生成器。
- [x] 视第一批拆分结果更新 `docs/architecture.md` 的 Excel 组件说明。
- [x] 拆出 `MainWindow.xaml.cs` 的输入状态和 options 构造。
- [ ] 继续拆 `MainWindow.xaml.cs` 的 workflow 编排和省份 UI 状态应用，作为下一批代码质量目标。
- [ ] 重庆阶段二实机测试如发现问题，暂停代码质量线并优先处理实测问题。
