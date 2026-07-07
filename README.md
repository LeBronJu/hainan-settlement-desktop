# 多省份售电结算桌面工具

这是售电结算自动化工具的独立 C# 桌面版项目。项目从海南成熟流程起步，正在演进为可接入多个省份的结算自动化工具。

项目从原仓库 `hainan-settlement-tool` 的 `csharp/` 子目录拆出，用于后续独立开发、测试、打包和发布。Python 版仍保留在原仓库，作为完整功能基线。

GitHub 仓库：

- `https://github.com/LeBronJu/hainan-settlement-desktop`

## 当前版本

- 最新正式版本：`v1.0.1`
- Release 页面：`https://github.com/LeBronJu/hainan-settlement-desktop/releases/tag/v1.0.1`
- 当前主线：`main`
- `main` 已包含 `v1.0.1` 之后合入的员工电量奖励、Win10/11 WPF 主题支持，以及重庆阶段一电量数据清洗和台账更新；这些改动尚未打新 tag 或正式发版。
- 当前正式包：
  - `HainanSettlementTool-Win7-8-v1.0.1.zip`
  - `HainanSettlementTool-Win10-11-v1.0.1.zip`

## 当前范围

- 已搭建 Win7/8版（WinForms）/ Win10/11版（WPF）/ Core / Excel 分层项目结构。
- 阶段1：电量处理表或 `.xlsx/.xls/.csv` 原始明细 -> 待整理台账 -> JSON 报告。
- 阶段1补充功能：只清洗电量数据，输出 `零售侧用户电量数据处理表.xlsx`。
- 海南阶段2：人工整理后的台账 -> 代理/居间分表、汇总表、JSON 报告、阶段二校验报告。
- 员工电量奖励：最新售电结算台账 + 月份范围 -> 员工电量奖励总表、每个负责人一份个人电量确认表、JSON 校验报告。
- 重庆阶段一：交易中心售电公司电量确认结算单 -> 重庆零售侧用户电量处理表、户号明细、重庆台账更新副本、JSON 校验/更新报告。
- 界面顶部“公共设置”里的结算月份和结果输出文件夹由阶段1、阶段2共用；员工电量奖励使用同一个结果输出文件夹，并在模块内选择开始/结束月份。
- Win10/11版（WPF）支持结算省份选择，启动时默认不选省份，需手动选择后再执行。未选省份时只显示空状态，不展示历史输入路径、执行按钮或省份专用输出项。海南显示成熟结算流程；重庆当前开放阶段一的“只清洗电量数据”和“清洗并更新台账”。
- Win10/11版（WPF）是主推荐版本和后续新功能入口。
- Win10/11版（WPF）支持界面主题选择：跟随系统、浅色、深色。主题只影响桌面工具界面，不影响生成的 Excel 文件。
- Win10/11版（WPF）的业务确认、警告和错误提示使用项目内现代弹窗样式，不再使用系统 `MessageBox`。
- Win7/8版（WinForms）进入维护模式：保留兼容包、构建和阻塞性 bugfix，不再做新省份功能或体验优化。
- 两个入口共享 Core/Excel 业务逻辑；结算规则和 workbook 修复应优先落在共享层。

## 项目结构

```text
HainanSettlementTool.sln
src/
  HainanSettlementTool.WinForms/   # Win7/8维护版界面，只修 bug 和兼容性
  HainanSettlementTool.Wpf/        # Win10/11主线界面，承接后续新功能和体验优化
  HainanSettlementTool.Core/       # 业务模型、业务服务、接口
  HainanSettlementTool.Excel/      # ClosedXML / ExcelDataReader 文件读写实现
docs/
  README.md                        # 文档地图：当前准绳、专题文档、历史 dev note
  CHANGELOG.md                     # 高信号历史里程碑索引
  architecture.md                  # 分层和迁移边界
  hainan-stage2-current-behavior.md # 海南阶段二当前行为
  RELEASE_CHECKLIST.md             # 正式发版前检查清单
  dev-notes/                       # 鲁棒性、解耦和临时技术记录
CONTEXT.md                         # 领域术语、业务口径和架构约束
HANDOFF.md                         # 当前版本状态、验证记录和下一步
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

仓库保留过参数化 real smoke 脚本用于授权真实工作副本验收；该脚本目前需要随已重命名的服务/网关刷新后再作为可运行入口使用。脚本不会内置真实数据路径，输出应写到指定 `OutputRoot` 下的新目录：

```powershell
.\scripts\run_real_smoke.ps1 -Month 5 -RawDetailPath "<原始明细.xls>" -ExistingPowerPath "<已清洗电量表.xlsx>" -BaseLedgerPath "<基础台账.xlsx>" -ReviewedLedgerPath "<人工整理后的台账.xlsx>" -ProxyTemplateDirectory "<上月代理分表文件夹>" -IntermediaryTemplateDirectory "<上月居间分表文件夹>" -SummaryTemplatePath "<上月汇总表.xlsx>" -OutputRoot "<临时输出文件夹>"
```

当前解决方案已编译通过。`v1.0.1` 发布前验证包含 Debug 测试、Release 构建、打包脚本、构建脚本兼容性检查，以及授权真实阶段二输出只读对比。

## 发布打包

生成 Win7/8版 Release 测试包：

```powershell
.\scripts\package_release.ps1
```

生成 Win10/11版 Release 测试包：

```powershell
.\scripts\package_wpf_release.ps1
```

脚本会执行 Release 构建，并在 `dist/` 下生成一个干净目录和 `.zip` 压缩包。测试时请保留目录内所有 `.dll` 和 `.config` 文件，不要只单独复制 exe。

Win7/8 包是维护版兼容包。正式发布仍可携带该包，但新功能和体验改进默认只进入 Win10/11 包，除非用户明确要求 Win7/8 适配。

正式发布时，GitHub Release 附件使用稳定 ASCII 文件名：

- `HainanSettlementTool-Win7-8-v1.0.1.zip`
- `HainanSettlementTool-Win10-11-v1.0.1.zip`

## 重要限制

- 原始零售侧明细可直接选择 `.xlsx`、`.xls` 或 `.csv`；清洗后的电量处理表仍输出为 `.xlsx`。
- 如果同一客户在原始明细中出现多个不同户号，程序不会自动任选一个户号，而是留给阶段1报告提示人工补齐。
- 阶段2保存时会用 ClosedXML 写入公式缓存；未接入 Excel 自动化。
- 阶段2按稳定 3 月参考文件夹改为模板驱动生成：分表复制上月 sheet 后只写输入列，汇总表保留模板隐藏列、合并表头、空白和日期显示格式。
- 海南阶段2详细当前行为以 `docs/hainan-stage2-current-behavior.md` 为准；重庆阶段2仍在分析/实现中，不能直接套用海南列位、公式和支付方默认。
- 阶段2会在生成前提示关键变化，并在生成后写入校验报告；后续月份仍建议先用工作副本验收。
- 员工电量奖励只依赖最新台账和月份范围，不需要另选奖励模板；隐藏月份电量列会按表头识别。
- 员工电量奖励遇到缺负责人、客户编号重复、企业名称为空但有电量等严重台账问题会停止生成，要求先检查台账。
- 深色模式只作用于 Win10/11版工具界面；所有生成的 Excel 表格保持固定浅色、打印友好样式，不读取或跟随系统主题。
- 重庆电量清洗和台账更新使用 `兆瓦时`，输出主表按用户名称汇总，并保留户号明细 sheet；该口径不得与海南 `万千瓦时` 台账口径混用。
- 重庆台账更新按 `电力用户名称` 精确匹配，不依赖也不自动补齐 `电力用户编码`；户号只保留在清洗明细和报告中用于追溯。匹配异常、多户号、疑似名称别名和已有电量差异会在写入前弹窗确认；用户可在弹窗里把本次未匹配的电量表客户显式选择为新增客户到台账、本月不写入，或匹配到一个已有台账客户。同一个已有台账客户在一次预检中只能被选择一次，新增和不写入可以重复选择。该选择只影响当前生成副本，并写入 JSON 更新报告。
