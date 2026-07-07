# Employee Reward Module

Date: 2026-07-02

Status: current module note. Use this file for employee reward scope and rules; stable cross-module vocabulary remains in `CONTEXT.md`.

This note records the implementation decisions for the independent `员工电量奖励` module.

## Reference And Real Files

- Reference folder recorded by the user: `D:\Document\文件处理\稳定参考版海南结算\电量奖励参考`
- 2026-07-02 implementation validation inspected the then-current 2026-05 production ledger read-only: `C:\Users\juqx2\Desktop\2026海南\海南2026-5月代理费结算\【2026年海南售电结算台账 - 0629】.xlsx`
- The user authorized read-only inspection for this analysis. Do not modify files in the production or reference folders.

The reference workbooks define the desired output shape, but they are not runtime inputs. The implemented module generates the reward workbooks from the latest ledger and an internal layout so the user does not need to select templates each month.

## Scope

Inputs:

- Latest Hainan settlement ledger.
- Start month and end month.
- Output folder.

Outputs:

- Employee reward summary workbook, for example `2026年1-5月员工电量奖励-海南.xlsx`.
- One employee confirmation workbook per负责人, for example `张三-2026年1-5月员工电量确认表-海南.xlsx`.
- JSON validation/report file.

If an output file already exists, the generator creates a timestamped unique filename instead of overwriting it.

## Ledger Reading Rules

- The official ledger sheet is identified by `海南2026年售电结算台账` first, then by required row-2 headers as fallback. The first sheet must not be assumed to be the ledger.
- Fixed columns are found by row-2 headers: `用电企业编号`, `用电企业名称`, `履约开始月份`, `项目开发人`, `代理或自营`, `负责人`.
- Month total columns are found by row-1 month label plus row-2 `总实际电量（万千瓦时）`. Hidden columns are valid and must be read.
- Rows with no customer code, no customer name, and no负责人 are treated as helper/check rows and excluded from employee reward detail rows.

## Business Rules

- Aggregate by `负责人`, not by `项目开发人`.
- `项目开发人` is an agent/intermediary relationship under a负责人, not the employee.
- Rows without负责人, duplicate customer codes, and rows with empty企业名称 but selected-period power are serious ledger errors. Generation stops and asks the user to check the ledger.
- Ledger and output power unit is `万千瓦时`.
- Reward formula is `电量（万千瓦时） * 10000 * 0.0001`; numerically, reward amount equals the power value.
- Generated workbook formulas keep the total and reward calculations visible for auditability.

## Implementation Shape

- Core:
  - `EmployeeRewardOptions`
  - `EmployeeRewardLedgerRow`
  - `EmployeeRewardDetail`
  - `EmployeeRewardSummary`
  - `EmployeeRewardResult`
  - `EmployeeRewardService`
  - `IEmployeeRewardExcelGateway`
- Excel:
  - `EmployeeRewardGenerator`
  - `ClosedXmlSettlementExcelGateway` implements the employee reward gateway in addition to the existing stage 1/2 gateways.
- WPF:
  - Adds an `员工电量奖励` tab.
  - UI collects ledger path, start/end months, and uses the shared output folder.
  - No WinForms UI was added; WinForms remains maintenance-only.

## Validation

- Synthetic Core tests cover selected-month aggregation, blocking ledger errors, single-month ranges, and shared workflow summary lines.
- Synthetic Excel tests cover official-sheet/month-column ledger reading, hidden month columns, helper/check row exclusion, non-overwrite output naming, generated workbook formulas, and personal confirmation workbook creation.
- A temporary real-data smoke test was run against the 2026-05 production ledger. It used a temporary output directory and deleted the output after validation; no real production file was modified or committed.
