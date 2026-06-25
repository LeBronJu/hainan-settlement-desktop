# C# 架构说明

## 分层

### HainanSettlementTool.WinForms

负责 Win7/8 版桌面界面：

- 文件选择
- 参数输入
- 日志展示
- 调用 `Stage1Service` / `Stage2Service`

禁止在 UI 层写 Excel 读写、客户匹配、金额计算等业务逻辑。

### HainanSettlementTool.Wpf

负责 Win10/11 版桌面界面壳：

- 文件选择
- 参数输入
- 日志展示
- 进度展示
- 预检确认弹窗
- 调用 `Stage1Service` / `Stage2Service`

禁止在 UI 层写 Excel 读写、客户匹配、金额计算等业务逻辑。

### HainanSettlementTool.Core

负责业务模型和业务服务：

- `Stage1Options`
- `Stage1Report`
- `PowerRow`
- `Stage1Service`
- `IStage1ExcelGateway`
- `Stage2Options`
- `Stage2Report`
- `Stage2PreflightReport`
- `Stage2SettlementCalculator`
- `FileAccessGuard`

Core 层只表达“要做什么”，不关心 Excel 文件具体怎么读写。

### HainanSettlementTool.Excel

负责 Excel 文件实现：

- `PowerWorkbookReader`：读取/写入电量处理表
- `RawDetailRowReader`：统一读取 `.xlsx/.xls/.csv` 原始零售侧明细行
- `RawDetailReader`：将原始明细行投影为电量行
- `CustomerCodeReader`：将原始明细行投影为户号映射
- `LedgerStage1Updater`：复制基础台账并写入当月电量
- `Stage2SettlementGenerator`：生成阶段2分表、汇总表和校验报告
- `ClosedXmlStage1ExcelGateway`：组合以上组件，对 Core 暴露统一接口

### Tests

测试项目位于 `tests/`：

- `HainanSettlementTool.Core.Tests`：覆盖阶段2金额规则、台账/分表差异问题、文件访问保护。
- `HainanSettlementTool.Excel.Tests`：用合成 workbook / `.csv` / `.xlsx` 覆盖原始明细读取、户号冲突处理、阶段1台账写入和输出副本保护。

真实 workbook 不提交到仓库。必要时只能在用户明确授权后做本地 smoke，并删除临时输出。

## 阶段1边界

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

## 阶段2边界

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
- 生成前预检关键变化：新增代理/居间关系、新增分表客户、利润单价变化、税率变化、读取上月模板失败
- 复用上月分表作为模板生成当月分表
- 复用汇总表模板插入当月区块并写入当月值/必要公式
- 保留模板里的隐藏列、合并表头、空白单元格、日期显示格式和非当月公式
- 按 Python 版历史口径输出代理费/居间费统计
- 输出阶段二校验报告，记录预检变化和分表金额核对差异

阶段2不负责：

- 修改台账客户名称以匹配汇总表/收款账户名称
- 自动补负责人或项目开发人
- 解释 1月/2月历史异常数据为通用规则
- 使用 Excel 自动化强制重算公式（当前仅用 ClosedXML 在保存前写入公式缓存）

## 发布和版本

- `main` 是当前发布主线。
- `v1.0` 是第一个正式版本。
- Win7/8 和 Win10/11 是两个长期共存的桌面入口，不是两个业务分支。
- 两个入口共享 Core/Excel 层；修业务规则时应优先修共享层。

正式 Release 附件采用 ASCII 文件名，避免 GitHub 或下载工具处理中文文件名不一致：

- `HainanSettlementTool-Win7-8-v1.0.zip`
- `HainanSettlementTool-Win10-11-v1.0.zip`

## 文档维护

文档更新是工程工作的一部分。任何行为、业务口径、发布包、脚本、目录结构或阶段边界变化，都应同步检查并更新：

- `README.md`
- `HANDOFF.md`
- `CONTEXT.md`
- `docs/architecture.md`
- 必要时新增或更新 `docs/dev-notes/` 下的日期文档

## 迁移策略

1. 继续用真实工作副本或脱敏样例验收 v1.0 输出，不覆盖正式文件。
2. 优先把共同行为沉到 Core/Excel，避免 Win7/8 和 Win10/11 双写漂移。
3. 继续补齐异常提示、脱敏 fixture 和自动化回归测试。
4. 最后再考虑替换 Python 版或保留为 legacy。
