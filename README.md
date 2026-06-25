# 海南售电结算桌面工具

这是海南售电结算自动化工具的独立 C# 桌面版项目。

项目从原仓库 `hainan-settlement-tool` 的 `csharp/` 子目录拆出，用于后续独立开发、测试、打包和发布。Python 版仍保留在原仓库，作为完整功能基线。

GitHub 仓库：

- `https://github.com/LeBronJu/hainan-settlement-desktop`

## 当前版本

- 最新正式版本：`v1.0`
- Release 页面：`https://github.com/LeBronJu/hainan-settlement-desktop/releases/tag/v1.0`
- 当前主线：`main`
- 当前正式包：
  - `HainanSettlementTool-Win7-8-v1.0.zip`
  - `HainanSettlementTool-Win10-11-v1.0.zip`

## 当前范围

- 已搭建 Win7/8版（WinForms）/ Win10/11版（WPF）/ Core / Excel 分层项目结构。
- 阶段1：电量处理表或 `.xlsx/.xls/.csv` 原始明细 -> 待整理台账 -> JSON 报告。
- 阶段1补充功能：只清洗电量数据，输出 `零售侧用户电量数据处理表.xlsx`。
- 阶段2：人工整理后的台账 -> 代理/居间分表、汇总表、JSON 报告、阶段二校验报告。
- 界面顶部“公共设置”里的结算月份和结果输出文件夹由阶段1、阶段2共用。
- Win7/8版与 Win10/11版长期共存，共享 Core/Excel 业务逻辑，不是两个互相独立的业务实现。

## 项目结构

```text
HainanSettlementTool.sln
src/
  HainanSettlementTool.WinForms/   # Win7/8版界面，只负责输入、日志、调用服务
  HainanSettlementTool.Wpf/        # Win10/11版界面，只负责输入、日志、调用服务
  HainanSettlementTool.Core/       # 业务模型、业务服务、接口
  HainanSettlementTool.Excel/      # ClosedXML / ExcelDataReader 文件读写实现
docs/
  architecture.md                  # 分层和迁移边界
  dev-notes/                       # 鲁棒性、解耦和临时技术记录
CONTEXT.md                         # 领域术语、业务口径和架构约束
HANDOFF.md                         # 当前版本状态、验证记录和下一步
```

## 设计原则

- UI 不写业务规则。
- Core 不引用 ClosedXML、WinForms、文件格式细节。
- Excel 层只做文件读写和模板复制。
- 每个阶段都输出可审计报告。
- C# 结果必须用脱敏样例和 Python 版输出对照后，才能视为可替代。
- 代码、业务口径、打包、发布或流程变化必须同步更新文档；文档滞后视为未完成。
- 每轮改动结束前必须明确判断文档是否受影响；受影响就更新并说明文件，不受影响也要说明原因。
- 合并前按 `.github/PULL_REQUEST_TEMPLATE.md` 里的文档检查项确认。

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

当前解决方案已编译通过。`v1.0` 发布前验证包含 Debug 测试、Release 构建、打包脚本、构建脚本兼容性检查，以及授权真实 `.xls` 原始明细 smoke test。

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

正式发布时，GitHub Release 附件使用稳定 ASCII 文件名：

- `HainanSettlementTool-Win7-8-v1.0.zip`
- `HainanSettlementTool-Win10-11-v1.0.zip`

## 重要限制

- 原始零售侧明细可直接选择 `.xlsx`、`.xls` 或 `.csv`；清洗后的电量处理表仍输出为 `.xlsx`。
- 如果同一客户在原始明细中出现多个不同户号，程序不会自动任选一个户号，而是留给阶段1报告提示人工补齐。
- 阶段2保存时会用 ClosedXML 写入公式缓存；未接入 Excel 自动化。
- 阶段2按稳定 3 月参考文件夹改为模板驱动生成：分表复制上月 sheet 后只写输入列，汇总表保留模板隐藏列、合并表头、空白和日期显示格式。
- 阶段2会在生成前提示关键变化，并在生成后写入校验报告；后续月份仍建议先用工作副本验收。
