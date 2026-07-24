# C# 多省份架构说明

当前文档入口见 `docs/README.md`。本文件只记录分层、模块边界和迁移策略；业务口径以 `CONTEXT.md` 和省份专题 current-behavior 文档为准。

## 分层

### HainanSettlementTool.WinForms

冻结的 Win7/8 历史兼容桌面界面，保留既有海南流程代码：

- 文件选择
- 参数输入
- 日志展示
- 调用 `HainanStage1Service` / `HainanStage2Service`

禁止在 UI 层写 Excel 读写、客户匹配、金额计算等业务逻辑。

WinForms 入口已冻结，不再承接新省份功能、体验优化、常规 bugfix 或发布打包，除非用户明确重新开启 Win7/8 支持。

### HainanSettlementTool.Wpf

负责 Win10/11 主线桌面界面壳：

- 文件选择
- 参数输入
- 省份选择
- 省份 UI 能力和文案 profile
- 省份 UI 状态应用控制器
- 日志展示
- 进度、状态和结果展示控制器
- 完成页按本次输出能力显示“查看可读报告”，只打开各省已发布的离线 HTML
- 日志展示和保存入口控制器
- 现代弹窗入口控制器
- 文件/文件夹选择入口控制器
- 输入状态、路径持久化和阶段 options 构造控制器
- 海南阶段一 workflow 编排控制器
- 重庆阶段一 workflow 编排控制器
- 阶段二 workflow 编排控制器
- 海南员工电量奖励 workflow 编排控制器
- `SettlementWorkflow` 构造工厂
- 现代 WPF 确认/错误弹窗和预检确认弹窗
- 调用 `HainanStage1Service` / `HainanStage2Service` / `ProvinceStage1Service` / `ChongqingStage2Service` / `GuangdongStage2MonthPreparationService`

禁止在 UI 层写 Excel 读写、客户匹配、金额计算等业务逻辑。
WPF 不再新增系统 `MessageBox` 作为业务确认、警告或错误提示；文件/文件夹选择器可以继续使用合适的系统对话框。

WPF 入口是后续新省份功能和体验优化的默认 UI 目标。

### HainanSettlementTool.Core

负责业务模型和业务服务：

- `HainanStage1Options`
- `HainanStage1Report`
- `HainanStage1RowMatchReport`
- `HainanPowerCleanReport`
- `HainanPowerRow`
- `HainanLedgerLayout`
- `HainanStage1Service`
- `IHainanStage1ExcelGateway`
- `HainanStage2Options`
- `HainanStage2SummarySubjectDecision`
- `HainanStage2PaymentParties`
- `HainanStage2Report`
- `HainanStage2PreflightReport`
- `HainanStage2CheckIssue`
- `HainanStage2Service`
- `IHainanStage2ExcelGateway`
- `Stage2SettlementCalculator`
- `HainanStage2AuditIssueFactory`
- `SettlementWorkflow`
- `HainanStage2WorkflowPlan`
- `HainanStage2WorkflowResult`
- `HainanEmployeePowerRewardOptions`
- `HainanEmployeePowerRewardService`
- `IHainanEmployeePowerRewardExcelGateway`
- `FileAccessGuard`
- `ProvinceCode`
- `ProvinceDisplayNames`
- `ProvinceStage1CleanOptions`
- `ProvinceStage1CleanResult`
- `ProvinceStage1LedgerUpdateOptions`
- `ProvinceStage1LedgerUpdatePlan`
- `ProvinceStage1LedgerUpdateResult`
- `ProvinceStage1CustomerDecision`
- `ProvinceStage1CustomerMatch`
- `ProvinceStage1LedgerUpdateIssueKinds`
- `ProvinceStage1Service`
- `IProvinceStage1ExcelGateway`
- `ChongqingStage2Options`
- `ChongqingStage2PreflightReport`
- `ChongqingStage2Report`
- `ChongqingStage2CheckIssue`
- `ChongqingStage2Service`
- `IChongqingStage2ExcelGateway`
- `GuangdongStage2MonthPreparationOptions`
- `GuangdongStage2PreflightReport`
- `GuangdongStage2MonthPreparationReport`
- `GuangdongStage2MonthPreparationService`
- `IGuangdongStage2MonthPreparationExcelGateway`

Core 层只表达“要做什么”，不关心 Excel 文件具体怎么读写。`Stage2SettlementCalculator` 保留为跨海南/重庆可复用的阶段二金额计算、容差和金额格式化 Module；海南阶段二预检问题构建放在 `HainanStage2AuditIssueFactory`，避免共享计算器泄漏海南专属报告模型。`SettlementWorkflow` 负责海南阶段一、海南只清洗、海南阶段二、海南员工电量奖励、多省份阶段一清洗/台账更新，以及重庆阶段二 Core 合同的 workflow 状态和共享结果摘要。UI 仍负责确认弹窗、预检决策、进度展示、日志展示和错误弹窗。

阶段二预检与整包发布形成两个有深度的共享 Module：

- 预检 Module 由 `Stage2PreflightIssue`、`Stage2PreflightIssueKinds`、`Stage2PreflightPolicy`、`Stage2RelationshipParameterValidator` 和 `Stage2OpaqueText` 组成，只表达稳定 issue code、`Blocker / RequiredDecision / Review / Information` 处置等级、关系参数合同、完整文本比较、是否允许继续，以及支付方/分表模板必选决策的待选、已解决、无效、冲突和过期状态。
- 结算完整性 Module 以 `Stage2BatchWorkspace` 为事务边界，以 Excel 层 `Stage2InputFingerprint`、`Stage2ManagedOutputInspector`、`Stage2BatchIntegrityVerifier` 和两省 writer 的强校验为门禁，表达确认后输入/规划输出状态不变、staging/回滚/失败留存、正式路径映射、发布前不变量和整包发布结果。

阶段二调用链保持单向：

```text
MainWindowStage2WorkflowController
  -> SettlementWorkflow.Plan{Province}Stage2
  -> 省份 Generator.Analyze（无副作用只读，产出预检事实、签名和指纹）
  -> Workflow Plan（冻结本次预检签名/指纹）
  -> Stage2PreflightPolicy + 共用 Stage2PreflightWindow
  -> SettlementWorkflow.Complete（复核初次 plan 未失效）
  -> 省份 Generator.Generate（再次核对签名/指纹）
       -> Stage2BatchWorkspace（唯一 staging）
       -> 省份 Split/Summary/Report Writer
       -> 省份业务强校验 + Stage2BatchIntegrityVerifier
       -> Stage2BatchWorkspace.Publish（整批发布或回滚/失败留存）
```

WPF 共用 `Stage2PreflightPresentationAdapter` 和 `Stage2PreflightWindow`，按 `(费用类型, 主体)` 把多条技术预检事实聚合成一张业务卡，收集支付方与可选分表模板决定，并控制确认按钮。模板候选浏览器把 Adapter 给出的完整候选集与当前可见批次分离：候选先随机洗牌、每批最多 5 个且同轮不重复，支持上一批、换一批、重新打乱，以及按主体/负责人/文件名对完整候选集执行空格分词 AND 搜索；已选模板不因换批失效。完整文件路径默认折叠；UI 不读取 workbook，也不重新解释税率、主体或模板规则。模板候选全集和是否必选进入预检签名；操作员所选路径必须仍属于该冻结全集。

海南和重庆通过窄 Adapter 投影省份关系事实并消费共享结果；台账列位、月份块、单位、金额公式、模板/sheet/样式、汇总列位、支付方 sheet 和省份例外继续留在各自 Excel Module。`自营不产生代理/居间结算` 是共享业务语义，但哪个字段是自营标记、项目开发人是否只是内部归属、退补如何处理，仍由省份 Adapter 明确投影，不能在共享层用固定列位猜测。不要引入吸收两省全部差异的 `BaseStage2Generator`、弱类型大 options 或共享层省份 switch。详细规则和实施状态见 `docs/dev-notes/stage2-preflight-integrity-2026-07-22.md`。

多省份客户处理应逐步收敛到通用业务语言：`匹配已有台账客户`、`新增客户到台账`、`本月不写入`。该能力可以先作为 Core 模型和 WPF 预检交互的共享 seam，但省份是否启用、默认行为如何执行由省份模块决定。海南阶段一当前保持成熟自动新增客户流程；重庆阶段一优先使用显式客户处理决定，以支持改名客户和新增客户混合出现的场景。

### HainanSettlementTool.Excel

负责 Excel 文件实现：

- `HainanPowerWorkbookReader`：读取/写入电量处理表
- `HainanRawDetailRowReader`：统一读取 `.xlsx/.xls/.csv` 原始零售侧明细行
- `HainanRawDetailReader`：将原始明细行投影为电量行
- `HainanCustomerCodeReader`：将原始明细行投影为户号映射
- `HainanStage1LedgerUpdater`：复制基础台账并写入当月电量
- `HainanLedgerWorkbookUtil`：识别海南台账主表并生成海南阶段一默认台账输出名
- `HainanStage2SettlementGenerator`：海南阶段二入口编排和预检问题构建；`Analyze` 只读产出预检签名/输入指纹，`Generate` 必须核对本次确认的签名/指纹后才允许写入 staging
- `HainanStage2LedgerReader` / `HainanStage2TemplateIndex` / `HainanStage2SplitWorkbookWriter` / `HainanStage2SummaryWorkbookWriter` / `HainanStage2ReportWriter`：海南阶段二 Excel 侧内部组件，分别负责台账读取、模板索引、代理/居间分表写入、汇总表写入、自包含 HTML/JSON/校验报告输出
- `HainanEmployeePowerRewardGenerator`：读取海南员工电量奖励所需台账字段和月度电量列，生成奖励总表、个人确认表和 JSON 校验报告
- `ChongqingPowerCleanGenerator`：读取重庆交易中心电量确认结算单，生成重庆用户电量汇总、户号明细和 JSON 校验报告
- `ChongqingLedgerStage1Updater`：读取重庆台账和重庆电量确认结算单，预检匹配异常，应用本次客户处理决定，写入目标月份电量到台账副本，并生成 JSON 更新报告
- `ChongqingStage2SettlementGenerator`：重庆阶段二入口编排，调用台账读取、分表写入、汇总写入和报告输出组件；按 30 列月度块识别代理费、居间费、退补电费主体，执行四级预检、存量/新增主体支付方决策、签名/指纹复核和整包发布门禁
- `ChongqingStage2LedgerReader` / `ChongqingStage2SplitWorkbookWriter` / `ChongqingStage2SummaryWorkbookWriter` / `ChongqingStage2ReportWriter`：重庆阶段二 Excel 侧内部组件，分别负责台账读取、代理/居间/退补分表写入、汇总表和支付方月度 sheet 写入、自包含 HTML/JSON/校验报告输出
- `ReadableReportDocument` / `ReadableHtmlReportRenderer`：跨省可读报告的窄展示 Module，只提供结构化卡片/表格模型、HTML 转义、离线响应式样式和列数校验；各省 Report Writer 负责把自己的报告事实映射进模型，不在共享 renderer 中加入省份 switch、结算规则或发布状态推断
- `Stage2InputFingerprint`：把台账、模板树、影响行为的配置，以及本批规划旧分表/正式汇总表的存在和内容状态压成稳定指纹，供初次 plan、生成准备和锁内发布前比对
- `Stage2ManagedOutputInspector`：只读扫描省份受管分表根，区分本批规划同月重跑、非规划目标月残留、不可读 workbook 和路径冲突；不删除年度历史 workbook
- `Stage2BatchIntegrityVerifier`：只做跨省可证明一致的物理文件门禁（路径必须位于本批 staging、文件非空、workbook 可重开、分表路径唯一、报告存在）；主体、金额、税率、支付方和 sheet 结构仍由两省 writer 各自强校验
- `GuangdongStage2MonthPreparationGenerator`：广东分表月份初始化入口；只扫描代理/居间/退补 workbook、严格识别数字月份 sheet、写输出副本和报告，不读取台账或计算金额
- `GuangdongStage2WorkbookInspector` / `GuangdongStage2WorkbookWriter` / `GuangdongStage2ReportWriter`：分别负责广东 workbook 预检分类、目标月 sheet 业务及默认打开视图归一化、未处理原文件保留和自包含 HTML/JSON/TXT 报告
- `IProvinceStage1Adapter`：Excel 层内部的多省份阶段一 adapter seam
- `ChongqingProvinceStage1Adapter`：组合重庆电量清洗和重庆台账更新实现
- `ClosedXmlUtil`：保留跨海南/重庆/广东共享的 ClosedXML 数字读取、列号转换和“目标工作表可见且唯一活动/选中”视图状态 helper，不包含省份台账表头或业务列位
- `ClosedXmlSettlementExcelGateway`：组合以上组件，对 Core 暴露统一接口，并通过省份 adapter 分发多省份阶段一 Excel 实现

广东阶段一实现时应新增独立 `GuangdongProvinceStage1Adapter`，由它组合广东 source reader、clean generator 和 ledger updater。现有 `ProvinceStage1Service` 仍有“只支持重庆”的临时门禁，`MainWindowChongqingStage1WorkflowController` 和共享摘要也仍含重庆/户号文案；接入广东前应把能力判断和展示中性化到 adapter/profile seam，而不是再增加一组省份 `if`。海南成熟阶段一继续走自己的 controller，不在这次扩展中强行迁移。

### Excel 引擎与批处理性能

当前生产 workbook 读写仍以 `ClosedXML 0.104.2` 为主。ClosedXML 已传递引入 `DocumentFormat.OpenXml 3.1.1`；Open XML SDK 是可选的底层 Adapter，不是已完成的全项目迁移。完整研究见 `docs/dev-notes/excel-performance-openxml-research-2026-07-10.md`。

批处理管线的架构原则：

- 先测量扫描、解析、业务计算、写入、公式评估和保存耗时，再选择优化手段。
- 预检生成不可变处理计划并冻结预检签名、输入/模板/配置和规划旧输出状态；确认后及发布前重新核对。当前安全性优先于性能，允许重复只读预检，后续只在等价性有测试保护后减少重复解析。
- 同一轮业务读取尽量只解析一次共享 workbook；为防止确认后输入或规划输出变化，允许在预检、生成准备和锁内发布前重复做只读文件指纹/哈希复核。同一 workbook 的建 sheet、写电量、金额和日期等修改仍集中到一次生成打开和一次保存。
- 并行只在独立 workbook 间实施，使用有上限的 worker 数并及时释放 workbook；不共享 ClosedXML/Open XML 可变文档实例。
- 性能管线只调度省份 Module，不吸收海南、重庆或广东的列位、单位、公式和模板规则。

Open XML 迁移顺序：

1. 当前暂停海南/重庆性能优化；如用户以后明确重启，先用海南建立性能基线和行为不变验收，不立即替换引擎。
2. 优先为台账读取、模板索引和其它只读投影建立 Open XML 影子 Adapter，与 ClosedXML 结果对比。
3. 局部写入只在公式缓存、shared strings、日期格式、隐藏结构和 WPS/Excel 兼容都有测试保护后切换。
4. `HainanStage1LedgerUpdater`、`HainanStage2SplitWorkbookWriter` 和 `HainanStage2SummaryWorkbookWriter` 等结构性 Writer 最后评估；如果收益不足，可长期保留 ClosedXML。
5. `.xls/.csv` 输入仍由现有 ExcelDataReader/CSV Adapter 负责，不纳入 Open XML 迁移。

### Tests

测试项目位于 `tests/`：

- `HainanSettlementTool.Core.Tests`：覆盖阶段2金额规则、台账/分表差异问题、文件访问保护、共享阶段 workflow 摘要、阶段二预检取消/继续流程、海南员工电量奖励汇总校验、多省份阶段一清洗摘要和重庆阶段一台账更新摘要。
- `HainanSettlementTool.Excel.Tests`：用合成 workbook / `.csv` / `.xlsx` 覆盖原始明细读取、户号冲突处理、阶段1台账写入、输出副本保护、阶段2分表新增客户、汇总表新增主体、模板页脚保护、海南员工电量奖励 workbook 读写、重庆电量清洗和阶段二写入，以及广东严格数字月份 sheet、目标月归一化、三类标题和重复运行。
- `HainanSettlementTool.Wpf.Tests`：覆盖阶段二同主体预检卡聚合、不同费用类型隔离、模板选项紧凑显示、完整路径折叠，以及模板候选 5 个分批、轮内不重复、重新打乱、选择保持和全量关键词 AND 搜索。

真实 workbook 不提交到仓库。必要时只能在用户明确授权后做本地 smoke，并删除临时输出。

## 海南阶段1边界

公共输入：

- 结算月份
- 结果输出文件夹（阶段1和阶段2共用）

输入：

- 基础台账
- 电量处理表，或 `.xlsx/.xls/.csv` 原始零售侧明细
- 可选参考台账

输出：

- 待整理台账
- 阶段1 JSON 报告

阶段1只负责：

- 写入当月电量
- 只清洗电量数据并输出 `零售侧用户电量数据处理表.xlsx`
- 新增客户名称
- 在原始明细户号唯一明确时补户号
- 输出待人工补齐项

阶段1不负责：

- 自动补负责人
- 自动判断代理/居间
- 在同一客户出现多个不同户号时自动选择其中一个
- 生成分表
- 生成汇总表

## 重庆阶段1边界

当前接入重庆阶段一的“只清洗电量数据”和“清洗并更新台账”：

输入：

- 重庆交易中心电量确认结算单（`.xlsx/.xls/.csv`）
- 重庆售电结算台账（台账更新时必填）
- 结算月份（可从文件标题/文件名识别，界面选择作为兜底）
- 输出文件夹

输出：

- `x月重庆零售侧用户电量数据处理表.xlsx`
- `x月重庆零售侧用户电量校验报告.json`
- `x月重庆售电结算台账-阶段一更新.xlsx`
- `x月重庆阶段一台账更新报告.json`

重庆阶段1已负责：

- 优先读取 `sheet1`，否则读取第一个 sheet
- 按表头识别 `用户名称`、`户号`、`时段`、`用电量`
- 将 `尖峰/高峰/平段/低谷` 映射为 `尖/峰/平/谷`
- 以 `兆瓦时` 为单位生成用户级汇总和户号明细
- 对缺关键字段、非法时段、非数字电量、负数电量停止生成
- 按重庆台账表头识别 `电力用户名称` 和目标月份 `总实际电量/尖/峰/平/谷`
- 目标月份区块不存在时，从上一月份重庆 30 列月度区块复制模板，清空目标月电量列后再写入
- 按 `电力用户名称` 精确匹配，并可接收 WPF 本次客户处理决定后写入台账副本，不自动补齐 `电力用户编码`
- 保留 `代理或自营=自营` 客户的代理/居间字段和月度收益列空白，不在阶段一补齐
- 对多户号、未匹配、新增、台账多出、疑似名称别名、已有电量差异进行写入前预检确认，并把本次客户处理决定写入 JSON 报告
- 对电量表客户不在台账的情况，WPF 预检应要求用户显式选择：新增客户到台账、匹配一个已有台账客户，或本月不写入；已有台账客户目标在同一次预检中只能被选择一次，新增/不写入动作可以重复选择

重庆阶段1暂不负责：

- 代理/居间结算生成
- 静默客户名称别名映射或模糊匹配
- 持久化客户别名表或跨月份自动复用人工匹配
- 自动维护 `电力用户编码` / B 列户号
- 读取或解释金额、电价字段

## 广东阶段1研究边界

广东阶段一当前仅完成结构和架构研究，尚未接入 Core、Excel 或 WPF。完整证据、预检建议和实施前待确认问题见 `docs/dev-notes/guangdong-stage1-research-2026-07-24.md`。

建议首版提供与重庆相似的“只清洗电量数据”和“清洗并更新台账”两种 workflow 形态，但不能复用重庆 workbook 实现。

输入契约：

- 只识别官方 `零售结算明细`、`市场联动价格`、`零售合同模式` 三个 sheet。
- 结算员后来制作的两个电量整理 sheet 完全排除，不是输入，也不能依赖其名称、顺序或结构。
- sheet 按名称、字段按表头识别；不能因为当前下载文件恰好把官方 sheet 放在前三位就用物理序号读取。
- 广东台账和当月明细用 `电力用户编码` 精确匹配，名称只作一致性检查。

建议 Excel 调用链：

```text
ProvinceStage1Service
  -> IProvinceStage1ExcelGateway
  -> ClosedXmlSettlementExcelGateway
  -> GuangdongProvinceStage1Adapter
       -> GuangdongStage1SourceReader
       -> GuangdongPowerCleanGenerator
       -> GuangdongLedgerStage1Updater
```

共享层只负责：

- options、plan、result、输入/输出安全和 workflow 生命周期；
- adapter 分发；
- WPF profile、进度、日志、通用预检外壳和结果展示；
- 通用结果合同与输出安全要求；具体 writer 和报告内容仍由广东 Adapter/Module 实现。

广东 Module 独立负责：

- 官方 sheet / 表头识别和人工 sheet 排除；
- 按客户编码聚合计量记录及兆瓦时单位；
- 广东 32 列月份块的识别、复制、清理和写入；
- 新客户插行、格式/公式模板和固定业务字段清空；
- 合同模式、市场联动价格和合同漂移的省份预检事实；
- 广东输出名和报告文案。

安全写入必须先清空所有台账数据行的目标月总量、峰、平、谷，再按编码写入聚合值；台账有但当月无电量的客户保持零。新增客户首版只写已确认安全的身份/电量，不从邻行继承合同、税务、代理参数或峰平谷系数。

广东编码身份还暴露出一个共享模型边界：现有 `ProvinceStage1CustomerDecision` 主要用名称表达来源和目标，不能直接承担广东可靠匹配。如广东需要人工“匹配已有”决定，模型应补稳定 source/target key；不要用名称冒充编码。

实现顺序建议：

1. 先写合成 workbook reader/聚合/预检测试。
2. 把 Core 支持门禁和 WPF 重庆专属文案中性化，但保持海南入口不变。
3. 注册广东 adapter，完成只清洗路径。
4. 在新客户插入和合同初始化规则确认后完成台账更新。
5. 用合成样例覆盖 32 列月份块、清零再写、新客户不继承业务值和重复运行。
6. 只有获得当次明确授权后，才用指定真实 workbook 做只读/副本 smoke。

## 海南阶段2边界

本节描述当前海南阶段二共享边界和可复用工程形态。海南阶段二细节见 `docs/hainan-stage2-current-behavior.md`；重庆阶段二设计见 `docs/dev-notes/chongqing-stage2-analysis-2026-07-07.md`。

输入：

- 人工整理后的当月台账
- 上月代理分表文件夹
- 上月居间分表文件夹
- 上月或修正版汇总表

输出：

- 当月代理分表
- 当月居间分表
- 当月代理费汇总表
- 阶段2 JSON 报告
- 自包含阶段2 HTML 可读报告和 TXT 校验/提示报告

阶段2负责：

- 按当月台账月份区块读取代理/居间明细
- 生成前预检关键变化：新增代理/居间关系、新增分表客户、上月分表存在本月台账外明细行、利润单价变化、税率变化、读取上月模板失败
- 新增汇总主体无法从历史汇总行继承支付方时，预检要求显式选择 `清能` 或 `清辉`，生成器拒绝缺少选择的运行
- 复用上月分表作为模板生成当月分表；新增主体借用同类型模板时，只在输出副本保留当前月份 sheet，并清除模板来源人的人工批注；精确命中存量主体时保留该主体自己的批注
- 复用汇总表模板插入当月区块并写入当月值/必要公式
- 保留模板里的隐藏列、合并表头、空白单元格、日期显示格式、非当月公式和表格格式
- 更新分表合计公式范围；分表底部已有签字日期时顺延一个月
- 更新汇总表合计公式范围；汇总表合计行之后的页脚/签字区不作为主体数据
- 按 Python 版历史口径输出代理费/居间费统计
- 输出自包含 HTML 可读报告，并保留 JSON/TXT 校验报告记录预检变化和分表金额核对差异

阶段2不负责：

- 修改台账客户名称以匹配汇总表/收款账户名称
- 自动补负责人或项目开发人
- 解释 1月/2月历史异常数据为通用规则
- 使用 Excel 自动化强制重算公式（当前仅用 ClosedXML 在保存前写入公式缓存）

## 重庆阶段2边界

当前已完成 Core 合同层、Excel workbook 写入和 Win10/11 WPF 生成入口。它会先执行预检和新增汇总主体支付方选择，确认后生成代理/居间/退补分表、汇总表副本、自包含 HTML 可读报告、JSON 报告和校验报告。该能力已通过合成测试、真实 5 月本地回放、授权 3-5 月阶段一+阶段二回测和用户实机核对，当前作为首个可用生产基线进入长期观察。

已完成：

- `ChongqingStage2Options` 表达台账、代理模板目录、可选居间模板目录、退补模板目录、汇总表模板、输出目录、预检签名/输入指纹，以及新增或存量缺失支付方主体和多份非精确同类分表模板的本次决策。
- `ChongqingStage2PreflightReport` / `ChongqingStage2CheckIssue` 表达重庆阶段二四级预检问题，包括关系资料、主体聚合、模板/目标月、收款人、税率、存量与新增支付方、受管输出和人工复核项。
- `ChongqingStage2Report` 表达代理费、居间费、退补电费三类结果摘要。
- `ChongqingStage2Service` 负责输入校验、预检授权校验、支付方决策校验和调用 Excel gateway；正式文件只由整包 staging/publish 流程发布。
- `SettlementWorkflow` 提供重庆阶段二 plan / complete / run 入口和摘要行，并在 Complete 前复核初次 plan 的签名/指纹仍然有效。
- `ClosedXmlSettlementExcelGateway` 已接入重庆阶段二 Excel gateway；`Analyze` 只读生成完整预检事实和签名/指纹，`Generate` 复核确认状态后写入 staging、强校验并整包发布 workbook 和报告。
- Win10/11 WPF 已开放重庆阶段二生成入口：输入台账、代理模板目录、可选居间模板目录、退补模板目录、汇总表模板和输出目录；确认预检后调用 `Generate`，写出分表、退补表和汇总表副本。

长期观察边界：

- 个别客户的临时人工调整不自动抽象为通用规则；发现可重复规则后再补预检或生成逻辑。
- 退补合计行之后的额外扣减块只同步已确认安全的 C-G 当月电量，H 列以后和汇总表抵扣/实际支付继续人工复核。
- 后续月份继续关注模板缺失、重复月份 sheet、日期文本和历史人工抵扣差异；这些情况应通过报告或预检显式暴露，不静默猜测。

## 广东分表月份初始化边界

当前广东入口是独立的低风险阶段二辅助 Module，不复用海南或重庆完整阶段二生成器，也不提前抽跨省通用阶段二抽象。

输入：

- 结算月份（当前支持 2-12 月）
- 至少一个代理、居间或退补分表根目录
- 输出目录

输出：

- 按代理/居间/退补和输入相对路径保存的正常 workbook 副本
- 跳过或生成失败时，按业务类型保存到 `【未处理-需人工复核】` 的输入原文件副本
- 广东分表月份初始化自包含 HTML 结果报告、JSON 报告和 TXT 校验报告

负责：

- WPF 省份 profile 控制广东只显示三个分表目录，阶段一、台账和汇总表输入隐藏。
- Core 验证输入目录、月份和输出安全边界，表达预检分类及生成结果。
- Excel 递归扫描 `.xlsx`，严格按数字 sheet 名选择标准月份，保留所有非标准 sheet。
- Excel 确保标准目标月可见，并把它设为正常输出的唯一活动且唯一选中 sheet；内容已经准备好但目标月隐藏或视图仍停在旧月份时，只在输出副本中做视图归一化。原输入不被修改；将归一化后的输出作为后续输入时进入原样复制路径。
- 目标月存在时保留目标月，目标月缺失时从标准上月复制；清理明细 C-F 并按确定规则更新日期。
- 跳过或生成失败时不修改业务内容，把输入原文件原样复制到对应类别的人工复核目录；原副本不新增目标月份 sheet。
- 每个 workbook 独立失败、整批继续；生成结果检查输入分类守恒，并在 HTML/JSON/TXT 报告和 WPF 完成状态中区分正常输出、已保留的跳过和失败。

不负责：

- 广东台账读取、电量写入、客户匹配、代理/居间/退补金额计算或汇总表生成。
- 把文件名、最后一个 sheet 或 `4-2月新增` 等特殊字符串当作月份来源判断。
- 覆盖输入 workbook，或对无法可靠识别、相互冲突的结构和日期做猜测或自动纠正。

## 海南员工电量奖励边界

输入：

- 最新海南售电结算台账
- 奖励开始月份和结束月份
- 输出文件夹

输出：

- 员工电量奖励总表
- 每个负责人一份员工电量确认表
- 员工电量奖励 JSON 校验报告

海南员工电量奖励负责：

- 按台账表头识别客户编号、企业名称、履约开始月份、项目开发人、代理/自营和负责人字段
- 按 `X月` + `总实际电量（万千瓦时）` 识别所选月份电量列，包括隐藏列
- 按负责人汇总客户电量和奖励金额
- 对缺负责人、客户编号重复、企业名称为空但有电量等严重台账错误停止生成
- 生成内置布局的奖励总表和个人确认表，不要求用户选择模板

海南员工电量奖励不负责：

- 代理费、居间费、税率、借支抵扣或支付方计算
- 自动补负责人或解释缺失负责人
- 读取项目开发人作为员工
- 覆盖既有输出文件
- 直接承担重庆或其它省份电量奖口径

## 发布边界

版本号、最新正式包和测试包路径由 `README.md` 与 `HANDOFF.md` 维护；正式发布流程由 `docs/RELEASE_CHECKLIST.md` 维护。本文件只保留架构相关发布原则：

- `main` 是当前发布主线。
- Win10/11 WPF 是主线桌面入口和唯一维护发布目标；Win7/8 WinForms 是冻结的历史兼容入口。
- 历史 WinForms 和 WPF 入口共享 Core/Excel 层；修业务规则时应优先修共享层，但不要为 WinForms 单独复制新 UI 逻辑。

## 文档维护

文档影响判断是工程工作的一部分。任何行为、业务口径、发布包、脚本、目录结构或阶段边界变化，都应同步检查是否影响：

- `README.md`
- `docs/README.md`
- `docs/CHANGELOG.md`
- `HANDOFF.md`
- `CONTEXT.md`
- `docs/architecture.md`
- 必要时新增或更新 `docs/dev-notes/` 下的日期文档

每轮改动结束前必须明确判断文档是否受影响；受影响才更新对应文档。本项目 PR 不强制；合并前如不走 PR，也要按 `.github/PULL_REQUEST_TEMPLATE.md` 的同等检查项手动确认文档影响和验证结果。

## 迁移策略

1. 继续把工具从海南成熟模块演进为多省份平台，新增省份功能默认从 WPF 入口接入。
2. 优先把结算行为沉到 Core/Excel；新 UI 功能默认只进入 Win10/11 WPF，Win7/8 WinForms 不再做维护适配。
3. 重庆阶段一已覆盖电量清洗和台账更新；后续扩展应继续保持阶段一边界，不要提前把重庆阶段二规则硬编码进共享层。
4. 海南/重庆性能优化和广东完整结算当前均暂缓；如性能工作明确重启，再从海南基线和只读 Adapter 开始。
5. 继续补齐异常提示、脱敏 fixture 和自动化回归测试。
6. 最后再考虑替换 Python 版或保留为 legacy。
