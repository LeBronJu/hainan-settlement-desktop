# 海南售电结算桌面工具

这是海南售电结算自动化工具的独立 C# 桌面版项目。

项目从原仓库 `hainan-settlement-tool` 的 `csharp/` 子目录拆出，用于后续独立开发、测试、打包和发布。Python 版仍保留在原仓库，作为完整功能基线。

GitHub 仓库：

- `https://github.com/LeBronJu/hainan-settlement-desktop`

## 当前范围

- 已搭建 WinForms / Core / Excel 三层项目结构。
- 阶段1：电量处理表或原始明细 -> 待整理台账 -> JSON 报告。
- 阶段2：人工整理后的台账 -> 代理/居间分表、汇总表、JSON 报告。
- 界面顶部“公共设置”里的结算月份和结果输出文件夹由阶段1、阶段2共用。

## 项目结构

```text
HainanSettlementTool.sln
src/
  HainanSettlementTool.WinForms/   # 桌面界面，只负责输入、日志、调用服务
  HainanSettlementTool.Core/       # 业务模型、业务服务、接口
  HainanSettlementTool.Excel/      # ClosedXML 文件读写实现
docs/
  architecture.md                  # 分层和迁移边界
```

## 设计原则

- UI 不写业务规则。
- Core 不引用 ClosedXML、WinForms、文件格式细节。
- Excel 层只做文件读写和模板复制。
- 每个阶段都输出可审计报告。
- C# 结果必须用脱敏样例和 Python 版输出对照后，才能视为可替代。

## 运行环境目标

- 开发：Visual Studio 2022 或具备 .NET Framework 4.7.2 targeting pack 的 MSBuild 环境。
- 运行：Windows 7 SP1 及以上，需安装 .NET Framework 4.7.2 或更高版本。

## 开发环境状态

当前开发机已经安装并验证：

- .NET SDK 8/9
- Visual Studio Build Tools 2022
- .NET Framework 4.7.2 targeting pack / SDK

推荐编译命令：

```powershell
& "C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe" ".\HainanSettlementTool.sln" /restore /p:Configuration=Debug /m
```

当前解决方案已编译通过，结果为 `0 个警告 / 0 个错误`。

## 发布打包

生成 Release 版测试包：

```powershell
.\scripts\package_release.ps1
```

脚本会执行 Release 构建，并在 `dist/` 下生成一个干净目录和 `.zip` 压缩包。测试时请保留目录内所有 `.dll` 和 `.config` 文件，不要只单独复制 exe。

## 重要限制

- C# 第一版暂不直接清洗 `.xls` 原始明细；请先另存为 `.xlsx`，或使用已清洗的电量处理表。
- 阶段2保存时会用 ClosedXML 写入公式缓存；未接入 Excel 自动化。
- 阶段2按稳定 3 月参考文件夹改为模板驱动生成：分表复制上月 sheet 后只写输入列，汇总表保留模板隐藏列、合并表头、空白和日期显示格式；后续月份仍建议先用工作副本验收。
