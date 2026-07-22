# 多省份售电结算桌面工具

这是售电结算自动化工具的独立 C# 桌面版项目。项目从海南成熟流程起步，正在演进为可接入多个省份的结算自动化工具。

项目从原仓库 `hainan-settlement-tool` 的 `csharp/` 子目录拆出，用于后续独立开发、测试、打包和发布。Python 版仍保留在原仓库，作为完整功能基线。

GitHub 仓库：

- `https://github.com/LeBronJu/retail-power-settlement-desktop`

## 当前版本

- 最新正式版本：`v1.2.0`
- Release 页面：`https://github.com/LeBronJu/retail-power-settlement-desktop/releases/tag/v1.2.0`
- 当前主线：`main`
- `v1.2.0` 重点加强海南/重庆阶段二预检和整批输出完整性，新增三省统一样式的可读 HTML 报告，并修复海南新增主体模板/借支状态及重庆自营行误阻断问题。
- 当前正式包：
  - `RetailPowerSettlementTool-Win10-11-v1.2.0.zip`

## 当前范围

- 已搭建 Win10/11版（WPF）/ Core / Excel 分层项目结构，并保留冻结的 Win7/8 历史 WinForms 入口。
- 阶段1：电量处理表或 `.xlsx/.xls/.csv` 原始明细 -> 待整理台账 -> JSON 报告。
- 阶段1补充功能：只清洗电量数据，输出 `零售侧用户电量数据处理表.xlsx`。
- 海南阶段2：人工整理后的台账 -> 代理/居间分表、汇总表、自包含 HTML 可读报告、JSON 报告、阶段二校验报告。
- 重庆阶段2：人工整理后的重庆台账 -> 代理/居间/退补分表、汇总表副本、自包含 HTML 可读报告、JSON 报告、阶段二校验报告。
- 员工电量奖励：最新售电结算台账 + 月份范围 -> 员工电量奖励总表、每个负责人一份个人电量确认表、JSON 校验报告。
- 重庆阶段一：交易中心售电公司电量确认结算单 -> 重庆零售侧用户电量处理表、户号明细、重庆台账更新副本、JSON 校验/更新报告。
- 广东分表月份初始化：批量处理代理/居间/退补 workbook，创建或整理标准目标月份 sheet，清空 C-F 电量并顺延日期，输出 workbook 副本和 HTML/JSON/TXT 检查报告；日期或结构异常的原文件会原样保留到 `【未处理-需人工复核】`，界面明确显示部分完成，不自动纠正业务日期。当前不计算广东结算金额。
- 广东第一批及后续安全加固包已通过用户实测；用户运行 600 多个真实 workbook 总耗时 4 分多钟，当前性能满足临时需求。2026-07-22 用户完成全省份实机测试，未发现新的问题。广东完整结算和海南/重庆性能优化仍暂缓，开发优先保障阶段二业务正确性与正式输出完整性。
- 界面顶部“公共设置”里的结算月份和结果输出文件夹由阶段1、阶段2共用；员工电量奖励使用同一个结果输出文件夹，并在模块内选择开始/结束月份。
- Win10/11版（WPF）支持结算省份选择，启动时默认不选省份，需手动选择后再执行。海南显示成熟结算流程；重庆开放阶段一和阶段二生成；广东当前只显示代理/居间/退补三个分表目录和“分表月份初始化”入口。切换省份时会清空上一省份专用输入，避免把其它省份路径误用于广东。
- Win10/11版（WPF）是主推荐版本和后续新功能入口。
- Win10/11版（WPF）支持界面主题选择：跟随系统、浅色、深色。主题只影响桌面工具界面，不影响生成的 Excel 文件。
- Win10/11版（WPF）的业务确认、警告和错误提示使用项目内现代弹窗样式，不再使用系统 `MessageBox`。
- Win7/8版（WinForms）已冻结为历史兼容入口，不再做常规维护、功能适配或新包发布；除非用户明确重新开启 Win7/8 支持。
- 历史 WinForms 入口和 WPF 入口仍共享 Core/Excel 业务逻辑；结算规则和 workbook 修复应优先落在共享层。

## 项目结构

```text
HainanSettlementTool.sln
src/
  HainanSettlementTool.WinForms/   # 冻结的 Win7/8 历史兼容界面，不再做常规维护
  HainanSettlementTool.Wpf/        # Win10/11主线界面，承接后续新功能和体验优化
  HainanSettlementTool.Core/       # 业务模型、业务服务、接口
  HainanSettlementTool.Excel/      # ClosedXML / ExcelDataReader 文件读写实现
docs/
  README.md                        # 文档地图：当前准绳、专题文档、历史 dev note
  CHANGELOG.md                     # 高信号历史里程碑索引
  architecture.md                  # 分层和迁移边界
  hainan-stage2-current-behavior.md # 海南阶段二当前行为
  dev-notes/excel-performance-openxml-research-2026-07-10.md # Excel 性能和引擎迁移研究
  RELEASE_CHECKLIST.md             # 正式发版前检查清单
  dev-notes/                       # 鲁棒性、解耦和临时技术记录
CONTEXT.md                         # 领域术语、业务口径和架构约束
HANDOFF.md                         # 当前版本状态、验证记录和下一步
dist/                              # Git 排除；仅放测试包、正式包和冻结的 Win7/8 包
local-validation/                  # Git 排除；仅放本地回测和 smoke 产物
```

## 设计原则

- UI 不写业务规则。
- 新省份功能优先使用省份模块命名和边界，不把重庆或未来省份口径塞进海南专用类。
- Core 不引用 ClosedXML、WinForms、文件格式细节。
- Excel 层只做文件读写和模板复制。
- 每个阶段都输出可审计报告。
- C# 结果必须用脱敏样例和 Python 版输出对照后，才能视为可替代。
- 代码、业务口径、打包、发布或流程变化必须做文档影响判断；受影响才更新对应文档，避免无关文档噪音。
- 每轮改动结束前必须明确判断文档是否受影响；受影响就更新并说明文件，不受影响也要说明原因。
- 本项目是本地单人开发，PR 不强制；如本地合并到 `main`，仍需按 `.github/PULL_REQUEST_TEMPLATE.md` 中同等文档检查项手动确认。

## 运行环境目标

- 开发：.NET SDK 8 或更高版本；也可使用能解析 SDK-style `.NET Framework 4.7.2` 项目的 Visual Studio / Build Tools。
- 运行：Windows 7 SP1 及以上，需安装 .NET Framework 4.7.2 或更高版本。

## 开发环境状态

推荐开发环境：

- .NET SDK 8 或更高版本
- .NET Framework 4.7.2 targeting pack / SDK，或 NuGet reference assemblies
- 可选：Visual Studio / Build Tools。打包脚本会优先使用 `dotnet msbuild`，找不到时再通过 `vswhere` 自动发现 MSBuild.exe

推荐编译命令：

```powershell
dotnet msbuild ".\HainanSettlementTool.sln" /restore /p:Configuration=Debug /m
```

如果当前终端 PATH 还没刷新，可使用完整路径：

```powershell
& "C:\Program Files\dotnet\dotnet.exe" msbuild ".\HainanSettlementTool.sln" /restore /p:Configuration=Debug /m
```

检查仓库中是否重新引入固定 Visual Studio / MSBuild 路径：

```powershell
.\scripts\check_build_portability.ps1
```

文档相关改动收工前运行：

```powershell
.\scripts\check_docs_guardrails.ps1
git diff --check
```

仓库保留过参数化 real smoke 脚本用于授权真实工作副本验收；该脚本目前需要随已重命名的服务/网关刷新后再作为可运行入口使用。脚本不会内置真实数据路径；显式 `OutputRoot` 优先，未指定时默认写入 Git 排除的 `local-validation/smoke/`：

```powershell
.\scripts\run_real_smoke.ps1 -Month 5 -RawDetailPath "<原始明细.xls>" -ExistingPowerPath "<已清洗电量表.xlsx>" -BaseLedgerPath "<基础台账.xlsx>" -ReviewedLedgerPath "<人工整理后的台账.xlsx>" -ProxyTemplateDirectory "<上月代理分表文件夹>" -IntermediaryTemplateDirectory "<上月居间分表文件夹>" -SummaryTemplatePath "<上月汇总表.xlsx>" -OutputRoot "<临时输出文件夹>"
```

当前解决方案已编译通过。2026-07-22 阶段二完整性、预检交互、借用模板批注清理、可读报告和重庆自营排除修复仅使用合成 workbook 做自动化验证：Core 80/80、Excel 126/126（海南阶段二 57/57、重庆阶段二 40/40）、WPF 展示/控件 4/4 通过，完整解决方案 Debug/Release 构建通过。用户随后完成全省份实机测试并确认未发现问题；上述改动随 Win10/11 WPF 正式版 `v1.2.0` 发布。

## 发布打包

生成 Win10/11版 Release 测试包：

```powershell
.\scripts\package_wpf_release.ps1
```

脚本会执行 Release 构建，并在 `dist/test-packages/` 下生成带时间戳的干净目录和 `.zip` 压缩包。已有同名目标时会停止而不覆盖；打包后会校验目录与 ZIP 的完整内容并输出 SHA-256。测试时请保留目录内所有 `.dll` 和 `.config` 文件，不要只单独复制 exe。

生成带稳定文件名的正式发布包：

```powershell
.\scripts\package_wpf_release.ps1 -ReleaseTag v1.2.0
```

带 `ReleaseTag` 的稳定文件名包默认写入 `dist/releases/`；这只是打包行为，正式 tag、GitHub Release 和发布仍需要单独明确授权。

Win7/8 打包脚本仅作为历史兼容脚本保留，默认目录为 `dist/legacy-win7-8/`；未来不再默认生成或发布 Win7/8 包，除非用户明确重新开启 Win7/8 支持。

未来正式发布时，GitHub Release 附件使用稳定 ASCII 文件名：

- `RetailPowerSettlementTool-Win10-11-<tag>.zip`

历史 `v1.0.1` Release 曾同时包含 Win7/8 和 Win10/11 两个资产；该事实不代表后续继续维护 Win7/8。

## 重要限制

- 原始零售侧明细可直接选择 `.xlsx`、`.xls` 或 `.csv`；清洗后的电量处理表仍输出为 `.xlsx`。
- 如果同一客户在原始明细中出现多个不同户号，程序不会自动任选一个户号，而是留给阶段1报告提示人工补齐。
- 阶段2保存时会用 ClosedXML 写入公式缓存；未接入 Excel 自动化。
- 阶段2按稳定 3 月参考文件夹改为模板驱动生成：分表复制上月 sheet 后只写输入列，汇总表保留模板隐藏列、合并表头、空白和日期显示格式。
- 海南阶段2详细当前行为以 `docs/hainan-stage2-current-behavior.md` 为准；重庆阶段2实现和长期观察边界见 `docs/dev-notes/chongqing-stage2-analysis-2026-07-07.md`，两省仍不能互套列位、公式、模板继承或支付方决策。
- 阶段2会在生成前提示关键变化，并在生成后写入校验报告；后续月份仍建议先用工作副本验收。
- 员工电量奖励只依赖最新台账和月份范围，不需要另选奖励模板；隐藏月份电量列会按表头识别。
- 员工电量奖励遇到缺负责人、客户编号重复、企业名称为空但有电量等严重台账问题会停止生成，要求先检查台账。
- 深色模式只作用于 Win10/11版工具界面；所有生成的 Excel 表格保持固定浅色、打印友好样式，不读取或跟随系统主题。
- 重庆电量清洗和台账更新使用 `兆瓦时`，输出主表按用户名称汇总，并保留户号明细 sheet；该口径不得与海南 `万千瓦时` 台账口径混用。
- 重庆台账更新按 `电力用户名称` 精确匹配，不依赖也不自动补齐 `电力用户编码`；户号只保留在清洗明细和报告中用于追溯。匹配异常、多户号、疑似名称别名和已有电量差异会在写入前弹窗确认；用户可在弹窗里把本次未匹配的电量表客户显式选择为新增客户到台账、本月不写入，或匹配到一个已有台账客户。同一个已有台账客户在一次预检中只能被选择一次，新增和不写入可以重复选择。该选择只影响当前生成副本，并写入 JSON 更新报告。
- 海南和重庆阶段二使用同一套预检处置语义和 WPF 确认窗口：同一 `(费用类型, 主体)` 的选择和复核项集中到一张主体卡，完整路径默认折叠。新主体或缺失支付方显式选择清能/清辉；海南新主体有多份可借用的同类型分表时，同卡显式选择一份模板，不再因“候选多于一份”直接阻断。新增主体借用其它主体模板时不带入来源人的历史 sheet 或批注；存量主体自己的批注继续保留。收款人按完整单元格文本继承，不拆分多人姓名；不生成付款路由 JSON。正式结果和 HTML/JSON/TXT 报告先在隔离区完整生成和核验，再整批发布；当前不提供部分汇总表。
- 重庆阶段二读取重庆台账、代理/居间/退补模板目录和汇总表，生成代理/居间/退补分表、汇总表副本、自包含 HTML 可读报告、JSON 报告和校验报告。首版已通过合成测试、授权真实 5 月本地回放 smoke 和 3-5 月阶段一+阶段二回测；退补额外扣减、历史人工抵扣等特殊项仍通过校验报告和完成后弹窗提示人工复核。
- 广东当前不是完整阶段二结算：只识别名称完全等于数字月份的标准 sheet；目标月存在时保留并整理，缺失时才从标准上月复制。非标准 sheet 原样保留；跳过或生成失败的输入原文件进入 `【未处理-需人工复核】`，不会新增目标月份 sheet。金额、台账和汇总表不处理。
