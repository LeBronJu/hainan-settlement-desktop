# 项目变更索引

本文件记录已经完成的高信号里程碑，避免把历史流水继续堆进 `HANDOFF.md`。它不是发布说明全文，也不是业务规格书；当前状态仍以 `README.md`、`HANDOFF.md`、`CONTEXT.md`、`docs/architecture.md` 和专题 current-behavior/dev-note 为准。

详细命令输出、旧 handoff 快照和中间验证细节可从 git history 追溯。

## 2026-07-10

- 重庆阶段二验收收口：用户确认当前数据验证基本没有问题；结合 3-5 月阶段一+阶段二真实回测结果，重庆阶段二从“待首次验收实现”调整为首个可用生产基线，后续月份继续长期观察。
- 重庆阶段二人工边界保持不变：个别客户临时调整不自动固化；退补额外扣减 H 列以后、汇总抵扣和实际支付继续通过报告及完成弹窗要求人工复核。

## 2026-07-09

- 重庆阶段二实机反馈修复：完成后复核弹窗改为数量/分类摘要；汇总表目标月 `清能X月` / `清辉X月` 移到 `汇总表` 后并合并首行标题；WPF 启动省份默认空白，重庆主 tab 改回 `代理费结算`，并预留禁用的 `员工电量奖励` tab。
- 重庆阶段二 3-5 月真实回测修复：目标月分表/payment-party sheet 已存在时保留目标月格式和隐藏列，不再删除后从上月复制；目标月缺失时仍从上月新建并刷新可见列。
- 重庆退补额外扣减块处理：可识别的额外块只自动同步 C-G 当月电量，H 列以后和汇总抵扣/实际支付保持人工复核边界，并在校验报告和 WPF 完成后弹窗提示。
- 重庆阶段二回测收敛：阶段一电量 3/4/5 月对比均为 0 差异；支付方月度 sheet 隐藏列差异清零；剩余 ReviewRequired 为日期/公式文本、退补 C 列公式写法和 3 月历史人工抵扣差异。
- 重庆一阶段+二阶段通用回测入口：`scripts/run_chongqing_backtest.ps1` 已作为当前主线真实回测工具，要求显式授权路径或 case JSON，输出到 `dist/`。
- 多省份代码质量第九批：WPF `MainWindowChongqingStage1WorkflowController` 接管重庆阶段一清洗和台账更新编排，主窗口继续瘦身。

## 2026-07-08

- 多省份代码质量第八批：完成全项目命名治理收口，海南阶段一 Core/Excel 合同、海南台账布局、海南原始明细读取器和海南阶段二明细行全部显式命名；`ClosedXmlUtil` 只保留跨省共享 helper，重庆阶段一 WPF 私有方法也改为重庆显式命名。
- 多省份代码质量第七批：海南阶段二 Core/WPF 合同显式改为 `HainanStage2...`，workflow 方法改为 `Analyze/Plan/Complete/RunHainanStage2`，并把海南台账/分表差异问题构建从共享 `Stage2SettlementCalculator` 移到 `HainanStage2AuditIssueFactory`。
- 多省份代码质量第六批：员工电量奖励 Core/Excel/WPF workflow 类型显式改为 `HainanEmployeePowerReward...`，模型字段收敛到 `ResponsiblePerson` / `ProjectDeveloper` / `MonthlyPowers`，并把模块 dev note 改为海南专属入口，为未来重庆电量奖预留并列实现空间。
- 多省份代码质量第五批：WPF `MainWindow.xaml.cs` 抽出 `MainWindowHainanStage1WorkflowController` 和 `MainWindowHainanEmployeePowerRewardWorkflowController`，集中海南阶段一写台账/清洗电量和员工奖励生成编排；重庆阶段一暂留主窗口等待实测反馈稳定。
- 多省份代码质量第四批：WPF `MainWindow.xaml.cs` 抽出 `MainWindowStage2WorkflowController` 和 `SettlementWorkflowFactory`，集中海南/重庆阶段二 plan-confirm-complete 编排和 workflow 装配；阶段二取消路径和预检决策回写保持不变。
- 多省份代码质量第三批：WPF `MainWindow.xaml.cs` 抽出 `MainWindowProvinceUiController`，集中结算月份启停、省份 tab/panel 可见性、省份文案和按钮启停；窗口继续保留事件入口和 workflow 编排。
- 多省份代码质量第二批：WPF `MainWindow.xaml.cs` 抽出 `MainWindowInputController`，集中路径载入/保存、已保存省份恢复、阶段 options 构造、月份/省份选择读取和输入清空；窗口继续保留事件入口和 workflow 编排。
- 多省份代码质量第一批：海南阶段二 Excel 生成器从通用 `Stage2SettlementGenerator` 拆分并显式命名为 `HainanStage2...` 组件，分离台账读取、模板索引、分表写入、汇总表写入和报告输出；行为目标保持不变。
- 重庆阶段二首版 workbook 写入：`ChongqingStage2SettlementGenerator` 拆分为入口编排、台账读取、分表写入、汇总写入和报告输出组件；支持代理/居间/退补分表、汇总表副本、JSON 报告和校验报告生成。
- 重庆阶段二 WPF 入口从“只预检”改为“预检确认后生成”，确认窗口和省份文案同步更新；Win7/8 WinForms 仍不做新功能维护。
- 重庆阶段二真实 5 月本地 smoke：使用授权重庆目录只读输入、输出到系统临时目录，生成代理 19 行/2 组、退补 4 行/3 组、居间 0 行/0 组，当前无金额差异和模板 warning。

## 2026-07-07

- Win7/8 支持冻结：WinForms 从维护兼容入口调整为历史兼容入口，未来默认只维护和打包 Win10/11 WPF，除非用户明确重新开启 Win7/8 支持。
- 重庆阶段二 WPF 预检入口：重庆省份 profile 开放阶段二预检，新增退补分表目录输入和重庆专用预检窗口；当前只执行 Analyze-only 预检和新增汇总主体支付方选择，不写出分表、退补表或汇总表。
- 重庆阶段二真实模板复查：确认当前授权目录内代理/退补 `.xlsx` 模板均为标准 xlsx 容器；`._*` / `~$*` 属于系统或锁文件噪音。真实 5 月 Analyze-only 预检可读取最新版汇总表和 20260512 历史模板，当前支付方新增问题为 0。
- 文档结构整理：新增 `docs/README.md` 作为文档地图，新增 `docs/hainan-stage2-current-behavior.md` 收敛海南阶段二当前行为；将 dev note 标记为当前任务、当前架构、当前流程、当前模块、当前政策或历史记录。
- Handoff 压缩：`HANDOFF.md` 从历史流水文档改为当前交接页，只保留分支状态、安全边界、最新包、当前验证摘要、重庆二阶段提醒和下一步。
- 文档守护脚本：新增 `scripts/check_docs_guardrails.ps1`，用于检查 `HANDOFF.md` 行数、重复长章节、文档地图必要路由、dev-note/current-behavior 状态行和历史真实数据授权措辞。
- 重庆阶段二分析：完成授权只读结构分析，形成 `docs/dev-notes/chongqing-stage2-analysis-2026-07-07.md`。关键结论包括重庆 30 列月度块、隐藏列保留、汇总表 `当年费用总计` 锚点、代理/居间/退补三类结算、代理少回收电能量电费、退补分段电价，以及新增汇总主体支付方需要预检显式选择。
- 重庆阶段二 Core 合同切片：新增重庆阶段二 options、preflight/report/issue 模型、service、Excel gateway interface 和 workflow 入口；当时还没有 Excel 侧实现或 WPF 入口。
- 海南阶段二新增汇总主体支付方守护：存量汇总主体继承路径不变；新增汇总主体支付方无法可靠继承时，预检要求显式选择 `清能`/`清辉`，缺少选择时拒绝生成，避免继续静默默认 `清辉`。
- 重庆阶段二 Excel Analyze-only 切片：新增 `ChongqingStage2SettlementGenerator` 的台账读取和预检，按重庆 30 列月度块识别代理费、居间费、退补电费主体，并对新增汇总主体要求支付方选择；workbook 生成仍明确未实现。
- WPF log-controller：合入低风险 WPF 解耦切片，新增 `MainWindowLogController` 管理日志追加、清空和保存；不改变结算计算、Excel 行为或用户可见日志文本。
- WPF log-controller 测试包：生成 `HainanSettlementTool-Win10-11-Release-20260707-150019.zip` 供本机测试；没有创建正式 tag 或 GitHub Release。
- 重庆阶段一客户处理决定和 UI 修复：WPF 预检要求未匹配客户显式选择新增、跳过或匹配既有台账客户；修复省份下拉和客户目标下拉显示内部类型名的问题；WPF 可见标题改为 `清能电力-结算自动化工具`。
- 重庆阶段一验收包：生成 `HainanSettlementTool-Win10-11-Release-20260707-140505.zip`，用户实机测试暂未发现阻塞问题。
- 多省份技术债切片：合入 WPF path picker、dialog、result/progress/log 等 controller 拆分；海南专用阶段入口、网关和省份显示名逐步中性化。

## 2026-07-06

- 重庆阶段一真实结构分析和实现推进：按授权只读工作簿结构实现重庆电量清洗和台账更新，保留户号明细，按 `电力用户名称` 匹配台账，写入目标月电量到输出副本。
- 重庆目标月份区块复制：当目标月区块缺失时，从上一月重庆 30 列月度块复制模板，改月标并清空目标月电量列，保留系数、退补、代理/居间公式和备注等模板列。

## 2026-07-02

- 员工电量奖励模块合入 `main`：Win10/11 WPF 增加独立 `员工电量奖励` 功能，按最新台账和月份范围生成奖励总表、个人确认表和 JSON 校验报告。
- 员工电量奖励实机测试：用户完成实际流程测试并反馈未发现阻塞问题。
- 员工电量奖励文档：`docs/dev-notes/hainan-employee-power-reward-module-2026-07-02.md` 记录海南模块范围、输入输出、读取规则和验证结论。

## 2026-06-30 到 2026-07-01

- 海南阶段二特殊明细行和借用模板清理合入：新增主体借用模板时不携带模板来源主体的历史月份 sheet；上月特殊明细行不自动继承到本月生成 sheet，并在预检中提示。
- 阶段二生成稳定性增强：补强分表合计行公式范围、合计行格式兜底、签字日期顺延、汇总表页脚保护等行为。

## 2026-06-29

- `v1.0.1` 正式发布：发布 Win7/8 和 Win10/11 两个 zip 资产，主要包含阶段二 workbook 模板修复。
- real smoke runner 设计：新增参数化真实工作副本 smoke 入口设计，要求路径由参数传入、不硬编码真实生产路径。后续服务/网关重命名后，该脚本需要刷新后再作为可运行入口。
- Win7/8 维护政策：确认 Win10/11 WPF 是新功能主线，Win7/8 WinForms 保持维护兼容入口。

## 2026-06-25

- 文档同步 gate：确立每轮开发都要做文档影响判断，但只更新职责真正受影响的文档。
- 鲁棒性和解耦优先级审查：记录阶段二、workflow、Excel 写入和 UI 拆分的后续改造方向；其中一部分已在后续切片完成。

## v1.0

- 第一个正式 C# 桌面版发布。
- 保持 Win7/8 WinForms 和 Win10/11 WPF 双入口包。
- 以海南阶段一、海南阶段二为成熟主路径，后续通过真实工作副本和合成测试逐步替代历史 Python 流程。
