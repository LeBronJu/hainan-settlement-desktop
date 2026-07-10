# 文档地图

本文件是当前文档入口。先看这里，再进入具体文档，避免把历史 dev note 当成最新规则。

## 当前准绳

| 文档 | 职责 | 什么时候读 |
| --- | --- | --- |
| `README.md` | 用户可见范围、运行环境、构建、打包、当前发布状态和重要限制。 | 用户安装、打包、发布、确认当前可用功能时。 |
| `HANDOFF.md` | 当前分支状态、最近验证结果、测试包路径、正在推进的任务和下一步。 | 每次接手当前线程或判断下一步时。 |
| `CONTEXT.md` | 稳定领域词汇、业务规则、数据安全约束。 | 改结算规则、阶段边界、客户/台账/电量口径时。 |
| `docs/architecture.md` | 当前分层、模块边界、迁移策略。 | 改 Core/Excel/UI seam、接入新省份、拆分模块时。 |
| `docs/hainan-stage2-current-behavior.md` | 海南阶段二当前实现行为和已验证规则。 | 参考海南阶段二、为重庆阶段二借鉴工程形态时。 |
| `docs/RELEASE_CHECKLIST.md` | 正式 tag / GitHub Release 前的验证和文档 gate。 | 正式发版、打 tag、上传 release assets 时。 |
| `docs/CHANGELOG.md` | 高信号历史里程碑索引。 | 需要了解项目演进历史，但不需要旧 handoff 全量流水时。 |
| `scripts/check_docs_guardrails.ps1` | 文档守护检查脚本。 | 文档相关改动收工前运行，防止 handoff 再膨胀或状态/路由缺失。 |

## 专题和任务说明

| 文档 | 状态 | 用法 |
| --- | --- | --- |
| `docs/dev-notes/chongqing-stage2-analysis-2026-07-07.md` | 当前模块 note | 重庆阶段二结构、已实现行为、回测结论和长期人工复核边界。 |
| `docs/dev-notes/guangdong-stage2-month-preparation-2026-07-10.md` | 当前模块 note | 广东代理/居间/退补分表标准月份 sheet 初始化规则、安全边界和验证结果。 |
| `docs/dev-notes/multi-province-code-quality-2026-07-08.md` | 当前任务 note | 多省份代码质量、海南命名中性化、大类拆分和执行连续性。 |
| `docs/dev-notes/multi-province-readiness-2026-07-07.md` | 当前架构 note | 新省份接入、WPF 省份 UI、Core/Excel 多省份 seam 工作前必须读。 |
| `docs/dev-notes/hainan-employee-power-reward-module-2026-07-02.md` | 当前模块 note | 海南员工电量奖励模块的输入、输出、读取规则和验证记录。 |
| `docs/dev-notes/win7-8-maintenance-mode-2026-06-29.md` | 当前政策 note | 判断 Win7/8 WinForms 冻结/退役边界，或用户明确重新开启 Win7/8 支持时。 |
| `docs/dev-notes/documentation-sync-gate-2026-06-25.md` | 当前流程 note | 判断文档影响、决定该更新哪些文档时。 |
| `docs/dev-notes/real-smoke-runner-2026-06-29.md` | 历史工具 note | 记录 real smoke runner 设计意图；当前脚本引用旧服务/网关名，刷新前不要当作可直接运行入口。 |
| `scripts/run_chongqing_backtest.ps1` | 当前回测工具 | 用户明确授权重庆真实数据路径后，参数化运行重庆阶段一+阶段二 1-N 月回测；输出到 `dist/` 或指定目录，不硬编码真实路径。 |
| `scripts/run_chongqing_stage2_backtest.ps1` | 历史/辅助回测工具 | 仅重庆阶段二的旧回测入口；需要只跑阶段二时可参考，主线回测优先用 `scripts/run_chongqing_backtest.ps1`。 |

## 历史记录

| 文档 | 状态 | 注意 |
| --- | --- | --- |
| `docs/dev-notes/robustness-architecture-priority-2026-06-25.md` | 历史架构审查和执行记录 | 里面的建议有些已完成。当前规则以 `CONTEXT.md`、`docs/architecture.md` 和本地图为准。 |

## Agent 工作流文档

| 文档 | 用法 |
| --- | --- |
| `docs/agents/domain.md` | 说明本仓库的领域文档入口，供 issue/PRD/triage 类技能读取。 |
| `docs/agents/issue-tracker.md` | GitHub Issues 作为 issue/PRD 跟踪位置的说明。 |
| `docs/agents/triage-labels.md` | triage 标签词汇表。 |

## 常见任务阅读顺序

| 任务 | 必读 |
| --- | --- |
| 每次接手或恢复线程 | `AGENTS.md`、`HANDOFF.md`、`docs/README.md` |
| 重庆阶段二实现 | `CONTEXT.md`、`docs/architecture.md`、`docs/hainan-stage2-current-behavior.md`、`docs/dev-notes/chongqing-stage2-analysis-2026-07-07.md`、`docs/dev-notes/multi-province-readiness-2026-07-07.md` |
| 新省份接入 | `CONTEXT.md`、`docs/architecture.md`、`docs/dev-notes/multi-province-readiness-2026-07-07.md` |
| 广东分表月份初始化 | `CONTEXT.md`、`docs/architecture.md`、`docs/dev-notes/guangdong-stage2-month-preparation-2026-07-10.md` |
| WPF UI 或交互改动 | `docs/architecture.md`、`docs/dev-notes/multi-province-readiness-2026-07-07.md`；涉及 WinForms 时再读 `docs/dev-notes/win7-8-maintenance-mode-2026-06-29.md` |
| Excel adapter / workbook 读写改动 | `CONTEXT.md`、`docs/architecture.md`，再读对应省份或模块的 current-behavior/dev-note |
| 海南阶段二行为判断 | `CONTEXT.md`、`docs/hainan-stage2-current-behavior.md` |
| 员工电量奖励改动 | `CONTEXT.md`、`docs/dev-notes/hainan-employee-power-reward-module-2026-07-02.md` |
| 发布、打包、tag 或 GitHub Release | `README.md`、`HANDOFF.md`、`docs/RELEASE_CHECKLIST.md` |
| 文档质量或文档同步规则改动 | `docs/README.md`、`docs/dev-notes/documentation-sync-gate-2026-06-25.md`、`HANDOFF.md` |
| 代码质量、多省命名或大类拆分 | `docs/architecture.md`、`docs/dev-notes/multi-province-readiness-2026-07-07.md`、`docs/dev-notes/multi-province-code-quality-2026-07-08.md` |
| issue、PRD 或 triage 工作 | `docs/agents/domain.md`、`docs/agents/issue-tracker.md`、`docs/agents/triage-labels.md` |
| 授权真实数据分析、回测或 smoke | `HANDOFF.md` 的数据安全边界、对应任务 dev-note；重庆一阶段+二阶段回测可用 `scripts/run_chongqing_backtest.ps1`；没有当前明确授权时先问用户 |

## 维护规则

- 稳定业务规则放 `CONTEXT.md`；专题细节可放 `docs/<topic>-current-behavior.md`，并从 `CONTEXT.md` 链过去。
- 当前架构边界放 `docs/architecture.md`；一次性审查、方案和任务分析放 `docs/dev-notes/`。
- 当前分支、测试包、最近验证和下一步只放 `HANDOFF.md`，不要把它当业务规格书或历史流水。
- 已完成的重要历史里程碑放 `docs/CHANGELOG.md`；详细命令输出、旧 handoff 快照和中间流水交给 git history。
- dev note 顶部必须写明状态。允许状态包括：当前任务 note、当前架构 note、当前流程 note、当前模块 note、当前政策 note、当前/历史工具 note、历史架构审查和执行记录。
- current-behavior 文档顶部必须写明状态：当前行为文档。
- `HANDOFF.md` 超过 250 行、重新出现 Useful Files / Documentation Map 这类长索引章节，或追加长 release 历史章节，视为文档结构退化；应拆到 `docs/CHANGELOG.md` 或 `docs/dev-notes/`。
- 文档相关改动收工前运行 `.\scripts\check_docs_guardrails.ps1` 和 `git diff --check`。
- 改代码、规则、发布、脚本或任务状态后，按 `docs/dev-notes/documentation-sync-gate-2026-06-25.md` 做文档影响判断。
