using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using HainanSettlementTool.Core.Models;
using HainanSettlementTool.Core.Services;

namespace HainanSettlementTool.Excel
{
    internal static class HainanStage2LedgerReader
    {
        internal static HainanStage2LedgerSnapshot ReadSnapshot(string ledgerPath, int month)
        {
            var snapshot = new HainanStage2LedgerSnapshot();
            using (var workbook = new XLWorkbook(ledgerPath))
            {
                var worksheet = HainanLedgerWorkbookUtil.MainSheet(workbook);
                var start = FindMonthStartColumn(worksheet, month);
                var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 1;

                for (var row = 4; row <= lastRow; row++)
                {
                    var customer = TextUtil.S(worksheet.Cell(row, 3).GetFormattedString());
                    if (string.IsNullOrWhiteSpace(customer))
                    {
                        continue;
                    }

                    var owner = TextUtil.S(worksheet.Cell(row, 10).GetFormattedString());
                    var developer = TextUtil.S(worksheet.Cell(row, 8).GetFormattedString());
                    var interName = TextUtil.S(worksheet.Cell(row, 19).GetFormattedString());

                    ReadRelationship(
                        worksheet,
                        row,
                        start,
                        customer,
                        owner,
                        interName,
                        "居间",
                        start + 7,
                        start + 8,
                        start + 10,
                        start + 12,
                        snapshot.IntermediaryRows,
                        snapshot.Relationships,
                        snapshot.Issues);
                    ReadRelationship(
                        worksheet,
                        row,
                        start,
                        customer,
                        owner,
                        developer,
                        "代理",
                        start + 13,
                        start + 14,
                        start + 16,
                        start + 18,
                        snapshot.ProxyRows,
                        snapshot.Relationships,
                        snapshot.Issues);
                }
            }

            BuildSubjectGroups(snapshot);
            return snapshot;
        }

        private static void ReadRelationship(
            IXLWorksheet worksheet,
            int ledgerRow,
            int start,
            string customer,
            string owner,
            string entity,
            string kind,
            int ratioColumn,
            int unitPriceColumn,
            int taxRateColumn,
            int cachedNetColumn,
            ICollection<HainanStage2DetailSettlementRow> rows,
            ICollection<HainanStage2RelationshipOccurrence> relationships,
            ICollection<HainanStage2CheckIssue> issues)
        {
            var ratio = ReadParameterCell(worksheet, ledgerRow, ratioColumn);
            var unitPrice = ReadParameterCell(worksheet, ledgerRow, unitPriceColumn);
            var taxRate = ReadParameterCell(worksheet, ledgerRow, taxRateColumn);
            var settlementKind = kind + "费";
            var validation = Stage2RelationshipParameterValidator.Validate(entity, ratio, unitPrice, taxRate);

            if (!validation.HasRelationship)
            {
                if (validation.Errors.Count > 0)
                {
                    issues.Add(new HainanStage2CheckIssue
                    {
                        Code = Stage2PreflightIssueKinds.RelationshipParametersWithoutSubject,
                        Disposition = Stage2PreflightDisposition.Blocker,
                        Severity = "错误",
                        Category = "关系主体为空但参数已填写",
                        Kind = settlementKind,
                        SettlementKind = settlementKind,
                        Customer = customer,
                        Owner = owner,
                        LedgerRow = ledgerRow,
                        CurrentValue = string.Join("、", validation.Errors.Select(error => error.ParameterName + "=" + error.DisplayValue)),
                        Message = "台账第" + ledgerRow + "行客户“" + customer + "”未填写" + kind + "主体，但已填写" + string.Join("、", validation.Errors.Select(error => error.ParameterName)) + "。",
                        Suggestion = "请补全" + kind + "主体，或清空该关系的比例、单价和税率后重新预检。"
                    });
                }

                return;
            }

            if (!validation.IsValid)
            {
                issues.Add(new HainanStage2CheckIssue
                {
                    Code = Stage2PreflightIssueKinds.RelationshipParametersInvalid,
                    Disposition = Stage2PreflightDisposition.Blocker,
                    Severity = "错误",
                    Category = "关系参数缺失或无效",
                    Kind = settlementKind,
                    SettlementKind = settlementKind,
                    Customer = customer,
                    Owner = owner,
                    Entity = entity,
                    LedgerRow = ledgerRow,
                    CurrentValue = string.Join("、", validation.Errors.Select(error => error.ParameterName + "=" + (string.IsNullOrWhiteSpace(error.DisplayValue) ? "空白" : error.DisplayValue))),
                    Message = "台账第" + ledgerRow + "行" + kind + "主体“" + entity + "”的" + string.Join("、", validation.Errors.Select(error => error.ParameterName)) + "为空、不是数字或不大于 0。",
                    Suggestion = "主体已填写时，比例、单价和税率必须全部填写为大于 0 的数字；请修正台账后重新预检。"
                });
                return;
            }

            relationships.Add(new HainanStage2RelationshipOccurrence
            {
                LedgerRow = ledgerRow,
                Customer = customer,
                Owner = owner,
                Entity = entity,
                Kind = kind,
                SettlementKind = settlementKind,
                TaxRate = taxRate.Value
            });

            var detail = CreateDetailRow(
                worksheet,
                ledgerRow,
                start,
                customer,
                owner,
                entity,
                kind,
                ratio.Value,
                unitPrice.Value,
                taxRate.Value,
                cachedNetColumn);
            if (detail.Total > 0 && HasSettlementAmount(detail))
            {
                rows.Add(detail);
            }
        }

        private static void BuildSubjectGroups(HainanStage2LedgerSnapshot snapshot)
        {
            var activeRows = snapshot.ProxyRows.Concat(snapshot.IntermediaryRows).ToList();
            foreach (var relationshipGroup in snapshot.Relationships
                .GroupBy(item => HainanStage2ExcelUtil.SummaryKey(item.Entity, item.SettlementKind))
                .OrderBy(group => group.Min(item => item.LedgerRow)))
            {
                var ordered = relationshipGroup.OrderBy(item => item.LedgerRow).ToList();
                var first = ordered[0];
                var owners = ordered
                    .Select(item => TextUtil.S(item.Owner))
                    .Where(owner => !string.IsNullOrWhiteSpace(owner))
                    .Distinct(StringComparer.Ordinal)
                    .ToList();
                var rows = activeRows
                    .Where(row => HainanStage2ExcelUtil.SummaryKey(row.Entity, row.Kind + "费") == relationshipGroup.Key)
                    .OrderBy(row => row.LedgerRow)
                    .ToList();
                var taxRates = ordered
                    .Select(item => item.TaxRate)
                    .GroupBy(HainanStage2ExcelUtil.NormalizeTaxRate)
                    .Select(group => group.First())
                    .ToList();

                if (string.IsNullOrWhiteSpace(first.Owner))
                {
                    snapshot.Issues.Add(new HainanStage2CheckIssue
                    {
                        Code = Stage2PreflightIssueKinds.FirstOwnerMissing,
                        Disposition = Stage2PreflightDisposition.Blocker,
                        Severity = "错误",
                        Category = "首个关系行缺少负责人",
                        Kind = first.SettlementKind,
                        SettlementKind = first.SettlementKind,
                        Customer = first.Customer,
                        Entity = first.Entity,
                        LedgerRow = first.LedgerRow,
                        Message = first.SettlementKind + "主体“" + first.Entity + "”在台账中首次出现的第" + first.LedgerRow + "行缺少负责人，无法确定分表归属。",
                        Suggestion = "请在台账首次关系行补全负责人后重新预检。"
                    });
                }

                if (owners.Count > 1)
                {
                    snapshot.Issues.Add(new HainanStage2CheckIssue
                    {
                        Code = Stage2PreflightIssueKinds.MultipleOwners,
                        Disposition = Stage2PreflightDisposition.Review,
                        Severity = "提示",
                        Category = "同一主体关联多个负责人",
                        Kind = first.SettlementKind,
                        SettlementKind = first.SettlementKind,
                        Owner = first.Owner,
                        Entity = first.Entity,
                        LedgerRow = first.LedgerRow,
                        CurrentValue = string.Join("、", owners),
                        Message = first.SettlementKind + "主体“" + first.Entity + "”关联多个负责人，将按台账首次出现顺序归到“" + first.Owner + "”名下并合并成一张分表。",
                        Suggestion = "请核对负责人归属；如无需调整，可以继续生成。"
                    });
                }

                if (taxRates.Count > 1)
                {
                    snapshot.Issues.Add(new HainanStage2CheckIssue
                    {
                        Code = Stage2PreflightIssueKinds.ConflictingTaxRates,
                        Disposition = Stage2PreflightDisposition.Blocker,
                        Severity = "错误",
                        Category = "同一汇总主体扣税率冲突",
                        Kind = first.SettlementKind,
                        SettlementKind = first.SettlementKind,
                        Owner = first.Owner,
                        Entity = first.Entity,
                        LedgerRow = first.LedgerRow,
                        CurrentValue = string.Join("、", taxRates.Select(value => value.ToString("0.####%", CultureInfo.CurrentCulture))),
                        Message = first.SettlementKind + "主体“" + first.Entity + "”存在多个扣税率，一个汇总行无法可靠表达。",
                        Suggestion = "请统一台账中该主体、该费用类型的税率后重新预检。"
                    });
                }

                if (rows.Count == 0)
                {
                    continue;
                }

                var subjectGroup = new HainanStage2SubjectGroup
                {
                    Kind = first.Kind,
                    SettlementKind = first.SettlementKind,
                    Entity = first.Entity,
                    Owner = first.Owner,
                    FirstLedgerRow = first.LedgerRow,
                    TaxRate = taxRates[0]
                };
                subjectGroup.Owners.AddRange(owners);
                subjectGroup.Rows.AddRange(rows);
                snapshot.SubjectGroups.Add(subjectGroup);
            }
        }

        private static HainanStage2DetailSettlementRow CreateDetailRow(
            IXLWorksheet worksheet,
            int ledgerRow,
            int start,
            string customer,
            string owner,
            string entity,
            string kind,
            double ratio,
            double unitPrice,
            double taxRate,
            int cachedNetColumn)
        {
            var total = HainanStage2ExcelUtil.GetNumeric(worksheet, ledgerRow, start);
            var ledgerNet = HainanStage2ExcelUtil.GetNumeric(worksheet, ledgerRow, cachedNetColumn);
            var amounts = Stage2SettlementCalculator.CalculateAmounts(total, ratio, unitPrice, taxRate);

            return new HainanStage2DetailSettlementRow
            {
                LedgerRow = ledgerRow,
                Customer = customer,
                Owner = owner,
                Entity = entity,
                Kind = kind,
                Total = total,
                Sharp = HainanStage2ExcelUtil.GetNumeric(worksheet, ledgerRow, start + 1),
                Peak = HainanStage2ExcelUtil.GetNumeric(worksheet, ledgerRow, start + 2),
                Flat = HainanStage2ExcelUtil.GetNumeric(worksheet, ledgerRow, start + 3),
                Valley = HainanStage2ExcelUtil.GetNumeric(worksheet, ledgerRow, start + 4),
                PeakFlat = HainanStage2ExcelUtil.GetNumeric(worksheet, ledgerRow, start + 5),
                ValleyFlat = HainanStage2ExcelUtil.GetNumeric(worksheet, ledgerRow, start + 6),
                Ratio = ratio,
                UnitPrice = unitPrice,
                TaxRate = taxRate,
                LedgerNet = ledgerNet,
                Gross = amounts.Gross,
                Adjustment = amounts.Adjustment,
                AdjustedGross = amounts.AdjustedGross,
                TaxAmount = amounts.TaxAmount,
                CalculatedNet = amounts.CalculatedNet,
                ExpectedNet = amounts.ExpectedNet
            };
        }

        private static Stage2RelationshipParameterValue ReadParameterCell(IXLWorksheet worksheet, int row, int column)
        {
            return ReadParameterCell(worksheet, row, column, 0);
        }

        private static Stage2RelationshipParameterValue ReadParameterCell(IXLWorksheet worksheet, int row, int column, int depth)
        {
            var cell = worksheet.Cell(row, column);
            if (cell.HasFormula && depth < 8)
            {
                var formula = TextUtil.S(cell.FormulaA1).Replace("$", string.Empty);
                var match = Regex.Match(formula, "^([A-Z]{1,3})(\\d+)$", RegexOptions.IgnoreCase);
                if (match.Success)
                {
                    var targetColumn = HainanStage2ExcelUtil.ColumnNumber(match.Groups[1].Value.ToUpperInvariant());
                    var targetRow = Convert.ToInt32(match.Groups[2].Value);
                    if (targetRow != row || targetColumn != column)
                    {
                        return ReadParameterCell(worksheet, targetRow, targetColumn, depth + 1);
                    }
                }
            }

            if (cell.IsEmpty())
            {
                return new Stage2RelationshipParameterValue { DisplayValue = "空白" };
            }

            if (cell.DataType == XLDataType.Number)
            {
                var number = cell.GetDouble();
                return NumericParameter(number, number.ToString("0.###############", CultureInfo.InvariantCulture));
            }

            var text = cell.GetFormattedString();
            if (string.IsNullOrWhiteSpace(text))
            {
                return new Stage2RelationshipParameterValue { DisplayValue = "空白" };
            }

            double value;
            var normalized = text.Replace(",", string.Empty).Trim();
            if (double.TryParse(normalized, NumberStyles.Any, CultureInfo.InvariantCulture, out value)
                || double.TryParse(normalized, NumberStyles.Any, CultureInfo.CurrentCulture, out value))
            {
                return NumericParameter(value, text.Trim());
            }

            return new Stage2RelationshipParameterValue
            {
                HasContent = true,
                DisplayValue = text.Trim()
            };
        }

        private static Stage2RelationshipParameterValue NumericParameter(double value, string displayValue)
        {
            return new Stage2RelationshipParameterValue
            {
                HasContent = true,
                IsNumeric = true,
                Value = value,
                DisplayValue = displayValue
            };
        }

        private static bool HasSettlementAmount(HainanStage2DetailSettlementRow row)
        {
            return Math.Abs(row.LedgerNet) > Stage2SettlementCalculator.AmountTolerance
                || Math.Abs(row.CalculatedNet) > Stage2SettlementCalculator.AmountTolerance;
        }

        private static int FindMonthStartColumn(IXLWorksheet worksheet, int month)
        {
            var lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? 1;
            for (var column = 1; column <= lastColumn; column++)
            {
                if (TextUtil.S(worksheet.Cell(1, column).GetFormattedString()) == month + "月")
                {
                    return column;
                }
            }

            throw new InvalidOperationException("未找到 " + month + "月 的台账区块。");
        }

    }
}
