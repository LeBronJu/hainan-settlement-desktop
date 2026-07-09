# 多省份代码质量主线（2026-07-08）

状态：当前任务 note。代码质量、海南命名中性化、大类拆分和多省份可维护性工作先读本文件；稳定架构规则仍以 `docs/architecture.md` 和 `docs/dev-notes/multi-province-readiness-2026-07-07.md` 为准。

## 当前主线

今日并行两条主线：

1. 重庆阶段二实机验收支援。用户当前测试包为 `HainanSettlementTool-Win10-11-Release-20260708-155953.zip`；如果反馈 bug、输出差异或需要重做的规则，优先级高于代码质量拆分。
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
| 1519 | `src/HainanSettlementTool.Wpf/MainWindow.xaml.cs` | 第二至九批已处理主要 workflow 编排。已拆日志、进度、结果、弹窗、路径选择、输入状态/options 构造、省份 UI 状态应用、海南阶段一、重庆阶段一、阶段二和海南员工奖励 workflow；主窗口当前主要保留事件路由、主题、省份切换和少量壳层协调。 |
| 1207 | `src/HainanSettlementTool.WinForms/MainForm.cs` | 冻结。只因共享层修复被动受益，不安排主动拆分。 |
| 874 | `src/HainanSettlementTool.Excel/ChongqingPowerCleanGenerator.cs` | P2。重庆阶段一清洗稳定后可拆读取、校验、汇总、报告写入。 |
| 804 | `tests/HainanSettlementTool.Excel.Tests/HainanStage2SettlementGeneratorTests.cs` | 已随第一批重命名为 `HainanStage2SettlementGeneratorTests.cs`；后续再按行为分组拆测试。 |
| 803 | `src/HainanSettlementTool.Excel/ChongqingLedgerStage1Updater.cs` | P2。可按预检、客户决定应用、月份块写入、报告输出拆分。 |
| 721 | `tests/HainanSettlementTool.Excel.Tests/ChongqingPowerCleanGeneratorTests.cs` | 跟随重庆清洗拆分后再整理。 |
| 682 | `src/HainanSettlementTool.Excel/HainanEmployeePowerRewardGenerator.cs` | P2。命名已明确为海南专属；后续可按台账读取、汇总计算、workbook 输出拆分。 |
| 568 | `src/HainanSettlementTool.Wpf/MainWindow.xaml` | P2。先不为拆而拆，等 workflow/input 控制器稳定后再看 XAML 资源和控件分组。 |
| 544 | `src/HainanSettlementTool.Excel/ChongqingStage2SummaryWorkbookWriter.cs` | 观察。重庆阶段二刚完成首版，先等实机验收反馈，再拆长期字段、月度块、支付方 sheet。 |

## 命名治理规则

- `Hainan...` / `Chongqing...` 用于省份专属业务规则和 Excel 结构。
- `Province...` / `Settlement...` 只用于真实跨省共享模型、workflow、能力 profile 或 UI 壳。
- 不做项目名、解决方案名、程序集名、根命名空间和发布包名的一次性大迁移。
- 优先改最误导维护者的内部类名、变量名和 helper 名，例如实际只处理海南阶段二的通用名。
- 每个命名切片必须配套 focused 测试或构建验证，避免 XAML、脚本路径、反射字符串或项目引用被纯重命名破坏。

最终目标是全项目命名规范化：省份专属 Module 必须显式带省份名，真实跨省共享 Module 才允许使用通用名。重庆是新接入模块，当前命名基本已规避这个问题；海南历史通用名已在第六至八批集中收口。项目名、解决方案名、程序集名、根命名空间和发布包名暂不纳入当前批次，因为这属于高风险迁移。

## 命名治理保留边界

本轮已完成的海南历史通用名收口：

- `HainanLedgerLayout`：海南台账主 sheet、26 列月度块和固定资料区。
- `HainanStage1Options` / `HainanStage1Report` / `HainanStage1RowMatchReport` / `HainanPowerCleanReport` / `HainanPowerRow`：海南阶段一 Core 合同。
- `HainanStage1LedgerUpdater`：海南阶段一台账副本写入。
- `HainanPowerWorkbookReader` / `HainanRawDetailReader` / `HainanRawDetailRowReader` / `HainanRawDetailRow` / `HainanCustomerCodeReader`：海南电量处理表和原始零售侧明细读取。
- `HainanStage2DetailSettlementRow`：海南阶段二代理/居间明细行。
- `HainanLedgerWorkbookUtil`：海南台账主表识别和阶段一默认输出台账名。

刻意保留的通用名：

- `ClosedXmlUtil`：仅保留 `CellNumber` 和 `ColumnLetter`，海南台账逻辑已迁出。
- `Stage2SettlementCalculator` / `Stage2SettlementAmounts` / `GroupSettlementTotal`：海南和重庆阶段二共同使用。
- `ProvinceStage1...` Core 合同：表示多省份阶段一 seam；当前只支持重庆是能力覆盖问题，不应改回重庆专属名。
- `SettlementWorkflow` / `StageWorkflowResult<TReport>` / `ClosedXmlSettlementExcelGateway` / `ProvinceUiProfile`：组合型或跨省壳层。
- WPF 员工奖励页签、控件名和历史输入快照字段中的 `Reward...`：用户可见能力名仍是“员工电量奖励”，实际业务实现已由 `HainanEmployeePowerReward...` 隔离。

仍不纳入本批：

- 根项目名、命名空间、发布包名和用户路径不做一次性大迁移。

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
- `dotnet msbuild src\HainanSettlementTool.Wpf\HainanSettlementTool.Wpf.csproj /restore /p:Configuration=Release /m` 通过。
- `dotnet msbuild src\HainanSettlementTool.Wpf\HainanSettlementTool.Wpf.csproj /restore /p:Configuration=Debug /m` 通过。
- `dotnet msbuild src\HainanSettlementTool.Wpf\HainanSettlementTool.Wpf.csproj /restore /p:Configuration=Release /m` 通过。

## 第二批进展

2026-07-08 已完成 WPF 主窗口输入状态第一轮拆分：

- 新增 `MainWindowInputController`，负责载入/保存用户输入路径、恢复已保存省份、构造海南阶段一/阶段二、重庆阶段一/阶段二和员工电量奖励 options、读取结算月份/奖励月份/当前省份，以及清空阶段输入。
- `MainWindow.xaml.cs` 不再直接拼装阶段 options 或直接序列化 `UserInputSnapshot`；窗口当时仍保留按钮事件、workflow 编排、进度/结果/弹窗协调和省份 UI 状态应用，省份 UI 状态应用已在第三批拆出。
- `MainWindow.xaml.cs` 从 1519 行降到 1324 行。本轮不改变 UI 文案、输入字段、持久化文件格式或运行流程；同时补上原本“保存省份但不恢复省份”的输入状态缺口，并保护 `InitializeComponent()` 期间 SelectionChanged 早触发时的空引用风险。

本轮验证：

- `dotnet msbuild src\HainanSettlementTool.Wpf\HainanSettlementTool.Wpf.csproj /restore /p:Configuration=Debug /m` 通过。
- `dotnet test tests\HainanSettlementTool.Core.Tests\HainanSettlementTool.Core.Tests.csproj --no-restore` 通过，26 个测试。
- `dotnet test tests\HainanSettlementTool.Excel.Tests\HainanSettlementTool.Excel.Tests.csproj --no-restore` 通过，27 个测试。
- `dotnet msbuild src\HainanSettlementTool.Wpf\HainanSettlementTool.Wpf.csproj /restore /p:Configuration=Release /m` 通过。

## 第三批进展

2026-07-08 已完成 WPF 主窗口省份 UI 状态应用拆分：

- 新增 `MainWindowProvinceUiController`，负责结算月份启停、省份 tab 和 panel 可见性、省份文案、重庆阶段二退补输入显示、按钮启停和员工奖励 tab 退回逻辑。
- `MainWindow.xaml.cs` 不再直接操作省份 UI 状态的二十多个控件；窗口只读取当前 `ProvinceUiProfile`，调用控制器应用状态，并刷新结果区可见性。
- `MainWindow.xaml.cs` 从 1324 行降到 1306 行。本轮不改变 UI 文案、可见性规则、按钮启停规则、输入字段或运行流程。

本轮验证：

- `dotnet msbuild src\HainanSettlementTool.Wpf\HainanSettlementTool.Wpf.csproj /restore /p:Configuration=Debug /m` 通过。
- `dotnet test tests\HainanSettlementTool.Core.Tests\HainanSettlementTool.Core.Tests.csproj --no-restore` 通过，26 个测试。
- `dotnet test tests\HainanSettlementTool.Excel.Tests\HainanSettlementTool.Excel.Tests.csproj --no-restore` 通过，27 个测试。
- `dotnet msbuild src\HainanSettlementTool.Wpf\HainanSettlementTool.Wpf.csproj /restore /p:Configuration=Release /m` 通过。
- `dotnet msbuild src\HainanSettlementTool.Wpf\HainanSettlementTool.Wpf.csproj /restore /p:Configuration=Release /m` 通过。
- 子代理只读 review 未发现初始化顺序、tab 切换、空引用或结果区刷新回归。

## 第四批进展

2026-07-08 已完成 WPF 主窗口阶段二 workflow 编排拆分：

- 新增 `MainWindowStage2WorkflowController`，负责海南/重庆阶段二 options 创建后的 plan-confirm-complete 顺序、预检确认窗口、取消路径、进度、日志和结果摘要。
- 新增 `SettlementWorkflowFactory`，集中 WPF 入口对 `ClosedXmlSettlementExcelGateway`、阶段服务和 `SettlementWorkflow` 的装配。
- `MainWindow.xaml.cs` 的阶段二按钮事件现在只调用 controller；阶段一、清洗和员工奖励暂时只改用 workflow factory，未进一步改变运行流程。
- `MainWindow.xaml.cs` 从 1306 行降到 1032 行。本轮不改变阶段二确认文案、日志文案、进度百分比、取消行为或结果摘要。

本轮验证：

- `dotnet msbuild src\HainanSettlementTool.Wpf\HainanSettlementTool.Wpf.csproj /restore /p:Configuration=Debug /m` 通过。
- `dotnet test tests\HainanSettlementTool.Core.Tests\HainanSettlementTool.Core.Tests.csproj --no-restore` 通过，26 个测试。
- `dotnet test tests\HainanSettlementTool.Excel.Tests\HainanSettlementTool.Excel.Tests.csproj --no-restore` 通过，27 个测试。
- `dotnet msbuild src\HainanSettlementTool.Wpf\HainanSettlementTool.Wpf.csproj /restore /p:Configuration=Release /m` 通过。
- 子代理只读 review 未发现 options 创建、保存/确认顺序、预检决策回写、取消路径、结果文案或 dispatcher 使用回归。

## 第五批进展

2026-07-08 已完成 WPF 主窗口海南阶段一和员工奖励 workflow 编排拆分：

- 新增 `MainWindowHainanStage1WorkflowController`，负责海南阶段一写台账和海南只清洗电量的确认、保存输入、运行进度、日志和结果摘要。
- 新增 `MainWindowHainanEmployeePowerRewardWorkflowController`，负责员工电量奖励的期间确认、保存输入、运行进度、日志和结果摘要。
- `MainWindowInputController` 新增 `PrepareHainanPowerCleanInput`，保留海南只清洗电量时先计算输出路径、写入 `PowerBox`、再保存输入的既有行为。
- `MainWindow.xaml.cs` 从 1032 行降到 852 行。本轮不改变海南阶段一、海南清洗或员工奖励的确认文案、日志文案、进度百分比、结果摘要或输出路径行为。
- 重庆阶段一 workflow 当时暂留 `MainWindow.xaml.cs`，避免在重庆阶段二核对期间扩大重庆相关改动面；用户后续实机测试未发现阻断性问题后，已在第九批拆出。

本轮验证：

- `dotnet msbuild src\HainanSettlementTool.Wpf\HainanSettlementTool.Wpf.csproj /restore /p:Configuration=Debug /m` 通过。
- `dotnet test tests\HainanSettlementTool.Core.Tests\HainanSettlementTool.Core.Tests.csproj --no-restore` 通过，26 个测试。
- `dotnet test tests\HainanSettlementTool.Excel.Tests\HainanSettlementTool.Excel.Tests.csproj --no-restore` 通过，27 个测试。
- `dotnet msbuild src\HainanSettlementTool.Wpf\HainanSettlementTool.Wpf.csproj /restore /p:Configuration=Release /m` 通过。
- 子代理只读 review 未发现海南阶段一、海南清洗、员工奖励的保存/确认顺序、进度日志、结果摘要或输出路径行为回归。

## 第六批进展

2026-07-08 已完成海南员工电量奖励命名治理：

- Core/Excel 类型从泛化 `EmployeeReward...` 改为 `HainanEmployeePowerReward...`，包括 options、ledger row、detail、summary、result、output、service、gateway interface 和 generator。
- WPF workflow class 改为 `MainWindowHainanEmployeePowerRewardWorkflowController`；WPF 可见 tab 和控件名仍保留通用“员工电量奖励”，避免 UI 能力命名被省份污染。
- 模型字段从 `Owner` / `Developer` / `MonthPowers` 收敛为 `ResponsiblePerson` / `ProjectDeveloper` / `MonthlyPowers`，让“按负责人汇总”的口径更直接。
- `SettlementWorkflow` 入口改为 `RunHainanEmployeePowerReward`，避免未来重庆电量奖接入时误用海南台账单位和输出规则。
- `docs/dev-notes/employee-reward-module-2026-07-02.md` 重命名为 `docs/dev-notes/hainan-employee-power-reward-module-2026-07-02.md`，文档地图同步指向海南专属模块 note。

本轮不改变员工电量奖励输出文件名、sheet 名、公式、确认文案、UI 可见文字或现有生成流程。

本轮验证：

- `dotnet test tests\HainanSettlementTool.Core.Tests\HainanSettlementTool.Core.Tests.csproj --no-restore` 通过，26 个测试。
- `dotnet test tests\HainanSettlementTool.Excel.Tests\HainanSettlementTool.Excel.Tests.csproj --no-restore` 通过，27 个测试。

## 第七批进展

2026-07-08 已完成海南阶段二 Core/WPF 命名治理：

- 海南阶段二 Core 合同从通用 `Stage2Options`、`Stage2Report`、`Stage2PreflightReport`、`Stage2CheckIssue`、`Stage2PaymentParties`、`Stage2SummarySubjectDecision`、`Stage2WorkflowPlan`、`Stage2WorkflowResult` 改为 `HainanStage2...`。
- WPF 海南阶段二预检窗口从 `Stage2PreflightWindow` 改为 `HainanStage2PreflightWindow`。
- `SettlementWorkflow` 的海南阶段二入口改为 `AnalyzeHainanStage2`、`PlanHainanStage2`、`CompleteHainanStage2`、`RunHainanStage2`，避免未来重庆/其它省份误用。
- `Stage2SettlementCalculator` 保留共享金额计算、容差和格式化；海南台账/分表差异问题构建移到 `HainanStage2AuditIssueFactory`。
- 组合 gateway 内部字段改为 `_hainanStage2Generator`，WinForms 冻结入口只做编译跟随改名，不增加新功能。

本轮不改变海南阶段二输出文件名、sheet 名、公式、支付方预检规则、预检文案或生成流程。

本轮验证：

- `dotnet test tests\HainanSettlementTool.Core.Tests\HainanSettlementTool.Core.Tests.csproj --no-restore` 通过，26 个测试。
- `dotnet test tests\HainanSettlementTool.Excel.Tests\HainanSettlementTool.Excel.Tests.csproj --no-restore` 通过，27 个测试。
- `dotnet msbuild src\HainanSettlementTool.Wpf\HainanSettlementTool.Wpf.csproj /restore /p:Configuration=Debug /m` 通过。
- `dotnet msbuild src\HainanSettlementTool.Wpf\HainanSettlementTool.Wpf.csproj /restore /p:Configuration=Release /m` 通过。

## 第八批进展

2026-07-08 已完成全项目命名治理收口：

- 海南阶段一 Core 合同改为 `HainanStage1Options`、`HainanStage1Report`、`HainanStage1RowMatchReport`、`HainanPowerCleanReport`、`HainanPowerRow`。
- 海南台账布局改为 `HainanLedgerLayout`，并同步海南阶段一、海南阶段二、海南员工电量奖励和测试引用。
- 海南阶段一 Excel 侧改为 `HainanStage1LedgerUpdater`、`HainanPowerWorkbookReader`、`HainanRawDetailReader`、`HainanRawDetailRowReader`、`HainanRawDetailRow`、`HainanCustomerCodeReader`。
- `ClosedXmlUtil` 只保留跨省共享的 `CellNumber` / `ColumnLetter`；海南台账主表识别和阶段一默认输出名迁入 `HainanLedgerWorkbookUtil`。
- 海南阶段二明细模型改为 `HainanStage2DetailSettlementRow`。
- WPF 重庆阶段一私有 workflow 方法改为 `RunChongqingStage1CleanPowerAsync`、`RunChongqingStage1LedgerUpdateAsync` 和 `ConfirmChongqingStage1LedgerUpdate`；未在本批抽出新 controller，避免把命名治理扩大成 UI 结构重构。
- 海南阶段一和海南原始明细测试类、测试方法、fixture helper 同步显式命名。

本轮不改变任何业务规则、输出文件名、sheet 名、公式、JSON 字段或用户可见文案。

本轮验证：

- `dotnet test tests\HainanSettlementTool.Core.Tests\HainanSettlementTool.Core.Tests.csproj --no-restore` 通过，26 个测试。
- `dotnet test tests\HainanSettlementTool.Excel.Tests\HainanSettlementTool.Excel.Tests.csproj --no-restore` 通过，27 个测试。
- `dotnet msbuild src\HainanSettlementTool.Wpf\HainanSettlementTool.Wpf.csproj /restore /p:Configuration=Debug /m` 通过。

## 第九批进展

2026-07-09 已完成 WPF 重庆阶段一 workflow 编排拆分：

- 新增 `MainWindowChongqingStage1WorkflowController`，负责重庆阶段一只清洗电量和清洗并更新台账的确认、保存输入、预检确认、客户处理决定回写、运行进度、日志和结果摘要。
- `MainWindow.xaml.cs` 不再承载 `RunChongqingStage1CleanPowerAsync`、`RunChongqingStage1LedgerUpdateAsync` 或重庆阶段一预检确认逻辑，只保留省份能力检查和按钮事件路由。
- 本轮不改变重庆阶段一输入字段、预检窗口、客户处理决定规则、确认文案、日志文案、进度百分比、输出文件名或生成流程。
- `MainWindow.xaml.cs` 从 852 行降到 637 行。

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
- [x] 拆出 `MainWindow.xaml.cs` 的省份 UI 状态应用。
- [x] 拆出 `MainWindow.xaml.cs` 的阶段二 workflow 编排。
- [x] 拆出 `MainWindow.xaml.cs` 的海南阶段一和员工奖励 workflow 编排。
- [x] 将员工电量奖励 Core/Excel/WPF workflow 命名明确为海南专属，为未来重庆电量奖预留并列实现空间。
- [x] 将海南阶段二 Core/WPF 合同和 workflow 方法命名明确为海南专属，并把海南预检问题构建从共享阶段二金额计算器移出。
- [x] 完成全项目命名治理收口：海南阶段一、海南台账布局、海南原始明细读取、海南阶段二明细行和重庆阶段一 WPF 私有方法已显式命名；真实共享名保留。
- [x] 视重庆实机核对反馈和当前风险，拆出重庆阶段一 workflow 编排。
- [ ] 重庆阶段二实机测试如发现问题，暂停代码质量线并优先处理实测问题。
