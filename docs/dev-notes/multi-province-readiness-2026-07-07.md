# 多省份接入架构自查（2026-07-07）

状态：当前架构 note。新省份接入、WPF 省份 UI、Core 多省份 workflow、Excel 多省份 adapter 相关工作前仍需阅读本文件；当前文档入口见 `docs/README.md`。

## 背景

海南是成熟省份模块；重庆已经接入阶段一电量清洗和台账更新。工具后续会继续接入其它省份，因此下一轮新增省份前要先控制省份分支扩散，避免每加一个省就在 WPF、Core、Excel 三处同时堆 `if Hainan / if Chongqing / if XXX`。

本自查只看代码结构和文档，不读取真实 Excel、真实台账或结算输出。

## 当前已经稳定的 seam

1. Core / Excel 分层已经成立。
   - Core 表达阶段选项、结果、预检计划和 workflow 摘要。
   - Excel 负责 workbook 读取、清洗、写入和报告文件生成。
   - UI 只负责文件选择、确认、进度、日志和错误展示。

2. 多省份阶段一已经有初始 Module。
   - Core: `ProvinceStage1Service`、`ProvinceStage1CleanOptions`、`ProvinceStage1LedgerUpdateOptions`、`ProvinceStage1LedgerUpdatePlan`、`ProvinceStage1LedgerUpdateResult`。
   - Excel: `ChongqingPowerCleanGenerator`、`ChongqingLedgerStage1Updater`。
   - WPF: 通过省份选择进入重庆阶段一。

3. Win10/11 WPF 是新省份默认入口。
   - Win7/8 WinForms 已冻结为历史兼容入口。
   - 新省份 UI 和体验优化不应再要求 WinForms 同步实现，也不默认生成 Win7/8 包。

## 执行状态摘要

本文件保留原始自查发现，也记录 2026-07-07 已完成的低风险预布局切片。阅读时不要把下方每个“问题”都当成仍未处理。

已完成第一版：

- WPF `ProvinceUiProfile` 集中省份 UI 能力和文案。
- Excel 内部 `IProvinceStage1Adapter` / `ChongqingProvinceStage1Adapter` seam。
- `ProvinceStage1LedgerUpdateIssue.Kind` 稳定 issue code。
- WPF 进度、结果、弹窗、路径选择、日志和输入状态 controller 拆分。
- 省份显示名、Hainan Stage 1/2 服务名、组合型 ClosedXML gateway 等低风险命名中性化。
- 重庆客户处理决定模型和 WPF 预检交互；海南阶段一保持原有自动新增流程。

仍待处理：

- `ProvinceStage1Service` 的“是否支持”和通用必填项仍偏重庆专用，第三个省接入前应继续收敛到省份能力 seam。
- `MainWindow.xaml.cs` 仍保留 workflow 编排和省份 UI 状态应用逻辑，后续可继续低风险拆分。
- 项目名、命名空间、程序集名和发布包名暂不做一次性大迁移。
- 完整 MVVM、持久化客户别名表、跨省通用阶段二抽象仍是 P2。

## 原始自查发现（含已解决项）

### 1. 省份能力和 UI 文案分散

涉及文件：

- `src/HainanSettlementTool.Wpf/MainWindow.xaml.cs`
- `src/HainanSettlementTool.Core/Models/ProvinceCode.cs`
- `src/HainanSettlementTool.Core/Services/ProvinceStage1Service.cs`
- `src/HainanSettlementTool.Excel/ClosedXmlSettlementExcelGateway.cs`

问题：

- 当前省份选择、可用功能、按钮文案、标签文案和结果区展示主要集中在 `MainWindow.xaml.cs` 的条件判断里。
- Core 里也硬编码“当前仅支持重庆阶段一电量清洗/台账更新”。
- Excel gateway 里用 `if (options.Province == ProvinceCode.Chongqing)` 分发。

风险：

- 第三个省接入时，需要同时修改 WPF、Core、Excel 多处判断。
- UI 文案变化容易和 Core/Excel 支持能力不一致。
- 省份是否支持阶段一清洗、台账更新、阶段二、员工奖励，目前没有统一 Interface。

建议：

- 第三个省接入前，新增一个省份能力/界面配置 Module，例如 `ProvinceFeatureProfile` 或 `ProvinceUiProfile`。
- 这个 Module 至少表达：
  - 省份显示名。
  - 是否支持阶段一电量清洗。
  - 是否支持阶段一台账更新。
  - 是否支持阶段二。
  - 是否支持员工电量奖励。
  - 阶段一输入标签、按钮文案、空状态/结果区文案。
- WPF 读取该 profile 来更新界面，不直接散落省份三元判断。

收益：

- Locality：新增省份的 UI 能力和文案集中到一个地方。
- Leverage：同一份 profile 同时驱动 tab 可见性、按钮启停、标签和结果区。

### 2. 多省份阶段一 Excel Adapter 分发还不够深

涉及文件：

- `src/HainanSettlementTool.Core/Services/IProvinceStage1ExcelGateway.cs`
- `src/HainanSettlementTool.Excel/ClosedXmlSettlementExcelGateway.cs`
- `src/HainanSettlementTool.Excel/ChongqingPowerCleanGenerator.cs`
- `src/HainanSettlementTool.Excel/ChongqingLedgerStage1Updater.cs`

问题：

- `IProvinceStage1ExcelGateway` 这个 seam 是对的，但当前 adapter 分发在 `ClosedXmlSettlementExcelGateway` 里写死重庆。
- 如果继续加省，gateway 会变成省份 switchboard，Interface 小但 Implementation 的省份知识会越来越多。

建议：

- 引入内部省份阶段一 Adapter seam，例如 `IProvinceStage1Adapter`。
- 每个省一个 adapter：
  - `ChongqingProvinceStage1Adapter`
  - 未来 `XxxProvinceStage1Adapter`
- `ClosedXmlSettlementExcelGateway` 只维护 `Dictionary<ProvinceCode, IProvinceStage1Adapter>` 并转发调用。

收益：

- Locality：某省 Excel 读取/写入/预检问题只在该省 adapter 内定位。
- Leverage：Core 和 WPF 不关心具体省份 Excel 实现。

### 3. `ProvinceStage1Service` 的验证规则仍是重庆专用

涉及文件：

- `src/HainanSettlementTool.Core/Services/ProvinceStage1Service.cs`
- `src/HainanSettlementTool.Core/Models/ProvinceStage1CleanOptions.cs`
- `src/HainanSettlementTool.Core/Models/ProvinceStage1LedgerUpdateOptions.cs`

问题：

- `ValidateCleanOptions` 和 `ValidateLedgerUpdateOptions` 当前直接拒绝非重庆省份。
- 输入字段现在适合重庆阶段一，但未来省份可能需要额外文件、不同月份来源、不同输出命名。

建议：

- 不要马上把 options 改成泛型字典。
- 第三个省接入前，先把“是否支持”和“通用必填项”放到省份 profile / adapter Interface。
- 若第三个省需要额外输入，再引入明确的 province-specific options 或 input descriptor，不用把所有省份揉成一个弱类型大对象。

收益：

- 保持 Interface 清晰，不为了未来假设牺牲当前可读性。
- 当第三个省真实需求出现时，再扩展有证据的 seam。

### 4. WPF `MainWindow.xaml.cs` 已接近省份 UI 分发上限

涉及文件：

- `src/HainanSettlementTool.Wpf/MainWindow.xaml`
- `src/HainanSettlementTool.Wpf/MainWindow.xaml.cs`
- `src/HainanSettlementTool.Wpf/MainWindowProgressController.cs`
- `src/HainanSettlementTool.Wpf/MainWindowResultController.cs`

问题：

- `UpdateProvinceUi` 同时处理 tab 标题、功能可见性、输入标签、按钮文案、结果区行可见性和提示文本。
- `MainWindow.xaml.cs` 还承载运行状态、进度条、五步状态和右侧结果卡片渲染，代码后置文件容易继续变成所有 UI 状态的集中点。
- 现在只有海南/重庆，复杂度可控；第三个省接入后会变浅：Interface 是一个窗口方法，Implementation 包含大量省份知识。

建议：

- 第三个省接入前先抽 WPF 省份 UI profile。
- 不必立刻做完整 MVVM。
- 先做到：
  - 静态布局仍留在 XAML。
  - 运行时变化文案来自 profile。
  - `UpdateProvinceUi` 只负责应用 profile 到控件。
  - 运行进度、状态 pill、步骤状态、右侧结果卡片等可复用 UI 状态先从 `MainWindow.xaml.cs` 拆到独立控制器，不急于完整 MVVM。

收益：

- Locality：省份 UI 调整集中。
- Leverage：新增省份时不需要反复理解整个 `MainWindow`。

### 5. 预检 Issue 现在靠中文 Category 字符串组织

涉及文件：

- `src/HainanSettlementTool.Core/Models/ProvinceStage1LedgerUpdateIssue.cs`
- `src/HainanSettlementTool.Wpf/ProvinceStage1LedgerPreflightWindow.xaml.cs`
- `src/HainanSettlementTool.Excel/ChongqingLedgerStage1Updater.cs`

问题：

- WPF 预检窗口用中文 `Category` 字符串判断哪些项目进“客户手动匹配”，哪些留在“其它预检项目”。
- 这能满足当前重庆，但对未来省份的预检动作来说不够稳。

建议：

- 下一个涉及预检动作的省份接入前，把 issue 增加稳定 code / kind，例如：
  - `PowerCustomerMissingInLedger`
  - `LedgerCustomerMissingInPower`
  - `PossibleAlias`
  - `MultiAccountCustomer`
  - `ExistingPowerDifference`
- UI 用 code / kind 判断行为，用中文标题只负责展示。

收益：

- 防止改中文文案导致 UI 行为变化。
- 支持未来不同省份复用同一种预检动作。

### 6. 测试还缺“省份接入合同”

涉及文件：

- `tests/HainanSettlementTool.Core.Tests/SettlementWorkflowTests.cs`
- `tests/HainanSettlementTool.Excel.Tests/ChongqingPowerCleanGeneratorTests.cs`

问题：

- 当前测试覆盖重庆清洗、重庆台账更新和 workflow 摘要。
- 但还没有一组“接入新省份必须满足什么”的合同测试。

建议：

- 第三个省接入时新增 checklist 式测试：
  - 省份 profile 可被 UI 读取。
  - 不支持的阶段明确禁用或报错。
  - 支持阶段一清洗的省份能通过 workflow 返回统一摘要。
  - 支持台账更新的省份有预检计划、确认后写副本、不覆盖源文件。
  - 不读取真实 Excel，只用合成 workbook 或脱敏 fixture。

收益：

- 新省份接入不再只靠人工记忆。
- 未来 agent 能按同一测试形状补省份。

### 7. 海南命名需要逐步中性化

涉及文件：

- `src/HainanSettlementTool.Core/`
- `src/HainanSettlementTool.Excel/`
- `src/HainanSettlementTool.Wpf/`
- `src/HainanSettlementTool.WinForms/`
- `tests/`

问题：

- 项目最早以海南工具为起点，项目名、命名空间和部分类名/变量名仍带有 `Hainan` 语义。
- 工具已经开始接入重庆，后续还会继续接入其它省份。如果所有新省份都挂在海南命名下，长期会影响理解和维护。

执行原则：

- 今日主线应纳入“海南命名逐步中性化”，但不能和结算 bugfix 混在同一个大改里。
- 暂不贸然重命名项目文件夹、解决方案名、程序集名、根命名空间和发布包名，避免制造大范围构建/打包/引用风险。
- 优先处理误导性的内部类名、变量名、方法名、文案分组名，例如把实际代表多省份能力的对象从海南语义改为 `Province` / `Settlement` / 具体省份名。
- 每个命名切片都要有可运行测试或构建验证，避免纯重命名引入 WPF XAML 绑定、项目引用、反射字符串或脚本路径问题。

收益：

- 让代码结构和“多省份混合工具”的产品方向一致。
- 降低后续接入第三个省份时的认知成本。
- 避免继续把重庆、未来省份的业务概念塞进海南专用命名里。

### 8. 客户处理决定需要抽成通用能力，但不强行改海南稳定流程

涉及文件：

- `src/HainanSettlementTool.Core/Models/`
- `src/HainanSettlementTool.Core/Services/ProvinceStage1Service.cs`
- `src/HainanSettlementTool.Excel/ChongqingLedgerStage1Updater.cs`
- `src/HainanSettlementTool.Wpf/ProvinceStage1LedgerPreflightWindow.xaml.cs`

问题：

- 重庆真实测试已经出现“改名客户”和“新增客户”同时存在的场景，单纯把未匹配客户视为新增或只允许匹配已有台账客户都不够。
- 海南阶段一目前稳定，且没有真实改名案例；为了抽象而改变海南交互会增加不必要的回归风险。

决策：

- 抽通用“客户处理决定”的业务语言是必要的，但先抽模型和接口语义，不先改变海南成熟输出行为。
- 通用动作至少包括：
  - `匹配已有台账客户`
  - `新增客户到台账`
  - `本月不写入`
- 重庆优先启用显式选择：每个待判断客户必须选择一个动作或目标。
- 海南当前继续沿用现有自动新增客户行为；未来如果出现海南改名/别名场景，再把同一套预检交互启用到海南。

UI 约束：

- 下拉第一项为 `新增客户到台账`，第二项为 `不匹配，本月不写入`，后续列出可匹配台账客户。
- `新增客户到台账` 和 `不匹配，本月不写入` 是动作，允许重复选择。
- 具体台账客户是写入目标，同一次预检中只能被一个电量客户选择，防止两个电量客户写入同一台账行。

收益：

- 既为多省份长期扩展建立统一模型，又避免扰动海南当前已验证流程。
- 重庆可以严谨处理改名、新增、跳过三种情况，用户不用在代码猜测和人工事后修正之间二选一。

## 建议执行顺序

### 2026-07-07 P0 进展

已完成八个低风险预布局切片：

1. WPF 新增 `ProvinceUiProfile`，集中海南/重庆的省份显示名、可用功能、阶段一输入文案、按钮文案、结果区文案和文件选择标题。`MainWindow` 仍负责控件赋值和事件入口，但不再把海南/重庆 UI 文案散落在 `UpdateProvinceUi` 里。
2. Excel 层新增内部 `IProvinceStage1Adapter` seam 和 `ChongqingProvinceStage1Adapter`。`ClosedXmlSettlementExcelGateway` 只按 `ProvinceCode` 分发多省份阶段一 Excel 实现，不再直接写重庆 if 分支。
3. `ProvinceStage1LedgerUpdateIssue` 新增稳定 `Kind`，中文 `Category` 保留给展示和 JSON 兼容。WPF 预检窗口用 `Kind` 判断客户手动匹配类 issue，并保留中文 `Category` fallback。
4. WPF 新增 `MainWindowProgressController`，把状态 pill、进度条、进度说明和五步状态渲染从 `MainWindow.xaml.cs` 拆出。主窗口暂保留薄 helper，避免一次性重写业务按钮流程。
5. WPF 新增 `MainWindowResultController`，把右侧完成卡片、输出项可见性、最近输出目录和各阶段成功摘要从 `MainWindow.xaml.cs` 拆出。主窗口仍保留 workflow 编排和打开文件夹按钮入口。
6. WPF 新增 `MainWindowDialogController` 和 `MainWindowPathPickerController`，把现代确认/错误弹窗入口以及文件/文件夹选择入口从 `MainWindow.xaml.cs` 拆出。
7. WPF 新增 `MainWindowLogController`，把运行日志追加、滚动、清空和保存文本文件行为从 `MainWindow.xaml.cs` 拆出。
8. WPF 新增 `MainWindowInputController`，把路径载入/保存、阶段 options 构造、月份/省份选择读取和输入清空行为从 `MainWindow.xaml.cs` 拆出。
9. Core/Excel/WPF 已实现重庆客户处理决定：`匹配已有台账客户`、`新增客户到台账`、`本月不写入`。海南阶段一保持原有稳定自动新增客户流程。

未完成且仍建议后续处理：

- 海南命名逐步中性化已进入今日多省份技术债主线。已完成第一批低风险切片：海南专用阶段入口改为 `HainanStage1Service` / `HainanStage2Service` 和对应 gateway interface；组合型 ClosedXML gateway 改为 `ClosedXmlSettlementExcelGateway`；省份显示名抽到 `ProvinceDisplayNames`；WPF 窗口标题、日志文件名和输入持久化目录改为中性名称并保留旧目录读取 fallback。暂不做项目名/命名空间/发布包名的一次性大重命名。
- 客户处理决定已作为多省份通用模型落地，并优先启用于重庆显式预检和台账写入。海南仍不启用手动匹配 UI，保持成熟自动新增客户流程；未来只有出现真实改名/跳过需求时再接入同一模型。
- `ProvinceStage1Service` 仍有重庆专用支持校验；第三个省接入前应继续把“是否支持”和通用必填项收敛到更明确的省份能力 seam。
- 完整 MVVM、持久化别名表和跨省通用阶段二抽象仍属于 P2，暂不展开。重庆阶段二本身已经是当前任务，应先做重庆专属实现和验证，不提前抽象成跨省通用阶段二。

### P0：第三个省接入前先做

1. 建立省份能力/界面 profile，先替换 WPF `UpdateProvinceUi` 里的省份文案和可见性判断。（已完成第一版）
2. 建立 Excel 内部省份阶段一 adapter seam，把重庆 adapter 从 `ClosedXmlSettlementExcelGateway` 的 if 分支里独立出来。（已完成第一版）
3. 给 `ProvinceStage1LedgerUpdateIssue` 增加稳定 code / kind，避免 UI 行为依赖中文分类。（已完成第一版）
4. 继续按低风险边界拆 `MainWindow.xaml.cs`：进度/状态、结果摘要、确认/错误弹窗入口、文件/文件夹选择入口、日志控制和输入状态/options 构造已拆出；后续应优先处理仍留在窗口里的 workflow 编排或省份 UI 状态应用逻辑。
5. 继续海南命名中性化的低风险切片：先改内部多省份概念的类名/变量名/helper 名；保留已经准确表达海南专属规则的名称，不做根项目名、程序集名、命名空间和发布包名的大迁移。

### P1：第三个省接入时同步做

1. 为新省份建立独立 Excel adapter，不把新省规则塞进重庆或海南类。
2. 补齐新省份 synthetic workbook 测试。
3. 更新 `CONTEXT.md` 的省份业务口径，并更新 `docs/architecture.md` 的阶段边界。

### P2：可以暂缓

1. 完整 MVVM 重构。
2. 多语言 `.resx` 资源化。
3. 持久化客户别名表。
4. 跨省通用阶段二抽象。

这些都需要更多真实省份需求后再做，当前不应为假设过度设计。

## 每次新省份工作前必须读

新省份接入、WPF 省份 UI、Core 多省份 workflow、Excel 多省份 adapter 相关改动前，必须阅读本文件，并用这里的 P0/P1/P2 顺序判断是否先还技术债。
