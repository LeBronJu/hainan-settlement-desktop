# C# 多省份架构说明

当前文档入口见 `docs/README.md`。本文件只记录分层、模块边界和迁移策略；业务口径以 `CONTEXT.md` 和省份专题 current-behavior 文档为准。

## 分层

### HainanSettlementTool.WinForms

负责 Win7/8 维护版桌面界面：

- 文件选择
- 参数输入
- 日志展示
- 调用 `HainanStage1Service` / `HainanStage2Service`

禁止在 UI 层写 Excel 读写、客户匹配、金额计算等业务逻辑。

WinForms 入口进入维护模式：保留构建、打包和阻塞性 bugfix，不再承接新省份功能或体验优化，除非用户明确要求。

### HainanSettlementTool.Wpf

负责 Win10/11 主线桌面界面壳：

- 文件选择
- 参数输入
- 省份选择
- 省份 UI 能力和文案 profile
- 日志展示
- 进度、状态和结果展示控制器
- 日志展示和保存入口控制器
- 现代弹窗入口控制器
- 文件/文件夹选择入口控制器
- 现代 WPF 确认/错误弹窗和预检确认弹窗
- 调用 `HainanStage1Service` / `HainanStage2Service` / `ProvinceStage1Service`

禁止在 UI 层写 Excel 读写、客户匹配、金额计算等业务逻辑。
WPF 不再新增系统 `MessageBox` 作为业务确认、警告或错误提示；文件/文件夹选择器可以继续使用合适的系统对话框。

WPF 入口是后续新省份功能和体验优化的默认 UI 目标。

### HainanSettlementTool.Core

负责业务模型和业务服务：

- `Stage1Options`
- `Stage1Report`
- `PowerRow`
- `HainanStage1Service`
- `IHainanStage1ExcelGateway`
- `Stage2Options`
- `Stage2SummarySubjectDecision`
- `Stage2PaymentParties`
- `Stage2Report`
- `Stage2PreflightReport`
- `HainanStage2Service`
- `IHainanStage2ExcelGateway`
- `Stage2SettlementCalculator`
- `SettlementWorkflow`
- `Stage2WorkflowPlan`
- `Stage2WorkflowResult`
- `EmployeeRewardOptions`
- `EmployeeRewardService`
- `IEmployeeRewardExcelGateway`
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

Core 层只表达“要做什么”，不关心 Excel 文件具体怎么读写。`SettlementWorkflow` 负责海南阶段一、海南只清洗、海南阶段二、员工电量奖励、多省份阶段一清洗/台账更新，以及重庆阶段二 Core 合同的 workflow 状态和共享结果摘要。UI 仍负责确认弹窗、进度展示、日志展示和错误弹窗。

多省份客户处理应逐步收敛到通用业务语言：`匹配已有台账客户`、`新增客户到台账`、`本月不写入`。该能力可以先作为 Core 模型和 WPF 预检交互的共享 seam，但省份是否启用、默认行为如何执行由省份模块决定。海南阶段一当前保持成熟自动新增客户流程；重庆阶段一优先使用显式客户处理决定，以支持改名客户和新增客户混合出现的场景。

### HainanSettlementTool.Excel

负责 Excel 文件实现：

- `PowerWorkbookReader`：读取/写入电量处理表
- `RawDetailRowReader`：统一读取 `.xlsx/.xls/.csv` 原始零售侧明细行
- `RawDetailReader`：将原始明细行投影为电量行
- `CustomerCodeReader`：将原始明细行投影为户号映射
- `LedgerStage1Updater`：复制基础台账并写入当月电量
- `Stage2SettlementGenerator`：生成阶段2分表、汇总表和校验报告；负责阶段2 workbook 模板复制、合计公式、模板格式兜底、签字日期写入，以及海南新增汇总主体支付方决策的预检和生成校验
- `EmployeeRewardGenerator`：读取员工电量奖励所需台账字段和月度电量列，生成奖励总表、个人确认表和 JSON 校验报告
- `ChongqingPowerCleanGenerator`：读取重庆交易中心电量确认结算单，生成重庆用户电量汇总、户号明细和 JSON 校验报告
- `ChongqingLedgerStage1Updater`：读取重庆台账和重庆电量确认结算单，预检匹配异常，应用本次客户处理决定，写入目标月份电量到台账副本，并生成 JSON 更新报告
- `ChongqingStage2SettlementGenerator`：当前只实现重庆阶段二台账读取和预检，按 30 列月度块识别代理费、居间费、退补电费主体，并对新增汇总主体要求支付方选择；分表、退补表和汇总表 workbook 写入尚未实现
- `IProvinceStage1Adapter`：Excel 层内部的多省份阶段一 adapter seam
- `ChongqingProvinceStage1Adapter`：组合重庆电量清洗和重庆台账更新实现
- `ClosedXmlSettlementExcelGateway`：组合以上组件，对 Core 暴露统一接口，并通过省份 adapter 分发多省份阶段一 Excel 实现

### Tests

测试项目位于 `tests/`：

- `HainanSettlementTool.Core.Tests`：覆盖阶段2金额规则、台账/分表差异问题、文件访问保护、共享阶段 workflow 摘要、阶段二预检取消/继续流程、员工电量奖励汇总校验、多省份阶段一清洗摘要和重庆阶段一台账更新摘要。
- `HainanSettlementTool.Excel.Tests`：用合成 workbook / `.csv` / `.xlsx` 覆盖原始明细读取、户号冲突处理、阶段1台账写入、输出副本保护、阶段2分表新增客户、汇总表新增主体、模板页脚保护、员工电量奖励 workbook 读写、重庆电量清洗 workbook 输出、重庆台账副本写入，以及重庆阶段二 Analyze-only 台账读取/新增汇总主体支付方预检。

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

阶段2负责：

- 按当月台账月份区块读取代理/居间明细
- 生成前预检关键变化：新增代理/居间关系、新增分表客户、上月分表存在本月台账外明细行、利润单价变化、税率变化、读取上月模板失败
- 新增汇总主体无法从历史汇总行继承支付方时，预检要求显式选择 `清能` 或 `清辉`，生成器拒绝缺少选择的运行
- 复用上月分表作为模板生成当月分表；新增主体借用同类型模板时，只在输出副本保留当前月份 sheet，不保留模板来源人的历史月份 sheet
- 复用汇总表模板插入当月区块并写入当月值/必要公式
- 保留模板里的隐藏列、合并表头、空白单元格、日期显示格式、非当月公式和表格格式
- 更新分表合计公式范围；分表底部已有签字日期时顺延一个月
- 更新汇总表合计公式范围；汇总表合计行之后的页脚/签字区不作为主体数据
- 按 Python 版历史口径输出代理费/居间费统计
- 输出阶段二校验报告，记录预检变化和分表金额核对差异

阶段2不负责：

- 修改台账客户名称以匹配汇总表/收款账户名称
- 自动补负责人或项目开发人
- 解释 1月/2月历史异常数据为通用规则
- 使用 Excel 自动化强制重算公式（当前仅用 ClosedXML 在保存前写入公式缓存）

## 重庆阶段2边界

当前完成 Core 合同层和 Excel Analyze-only 预检切片，尚未完成 workbook 生成或 WPF 入口；因此它不是用户可运行功能。

已完成：

- `ChongqingStage2Options` 表达台账、代理模板目录、可选居间模板目录、退补模板目录、汇总表模板、输出目录和新增汇总主体支付方决策。
- `ChongqingStage2PreflightReport` / `ChongqingStage2CheckIssue` 表达重庆阶段二预检问题，包含新增汇总主体支付方选择需求。
- `ChongqingStage2Report` 表达代理费、居间费、退补电费三类结果摘要。
- `ChongqingStage2Service` 负责输入校验、输出目录创建、支付方决策校验和调用 Excel gateway。
- `SettlementWorkflow` 提供重庆阶段二 plan / complete / run 入口和摘要行。
- `ClosedXmlSettlementExcelGateway` 已接入重庆阶段二 Analyze-only Excel gateway；`Generate` 仍明确未实现，避免误导为可生成结算结果。

仍未完成：

- `ChongqingStage2SettlementGenerator` 的分表、退补表和汇总表 workbook 写入。
- WPF 阶段二输入和预检支付方选择 UI。
- 真实 5 月回放验证。

## 员工电量奖励边界

输入：

- 最新海南售电结算台账
- 奖励开始月份和结束月份
- 输出文件夹

输出：

- 员工电量奖励总表
- 每个负责人一份员工电量确认表
- 员工电量奖励 JSON 校验报告

员工电量奖励负责：

- 按台账表头识别客户编号、企业名称、履约开始月份、项目开发人、代理/自营和负责人字段
- 按 `X月` + `总实际电量（万千瓦时）` 识别所选月份电量列，包括隐藏列
- 按负责人汇总客户电量和奖励金额
- 对缺负责人、客户编号重复、企业名称为空但有电量等严重台账错误停止生成
- 生成内置布局的奖励总表和个人确认表，不要求用户选择模板

员工电量奖励不负责：

- 代理费、居间费、税率、借支抵扣或支付方计算
- 自动补负责人或解释缺失负责人
- 读取项目开发人作为员工
- 覆盖既有输出文件

## 发布边界

版本号、最新正式包和测试包路径由 `README.md` 与 `HANDOFF.md` 维护；正式发布流程由 `docs/RELEASE_CHECKLIST.md` 维护。本文件只保留架构相关发布原则：

- `main` 是当前发布主线。
- Win10/11 WPF 是主线桌面入口；Win7/8 WinForms 是维护版兼容入口。
- 两个入口共享 Core/Excel 层；修业务规则时应优先修共享层。

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
2. 优先把结算行为沉到 Core/Excel；新 UI 功能默认只进入 Win10/11 WPF，Win7/8 WinForms 只做维护适配。
3. 重庆阶段一已覆盖电量清洗和台账更新；后续扩展应继续保持阶段一边界，不要提前把重庆阶段二规则硬编码进共享层。
4. 继续补齐异常提示、脱敏 fixture 和自动化回归测试。
5. 最后再考虑替换 Python 版或保留为 legacy。
