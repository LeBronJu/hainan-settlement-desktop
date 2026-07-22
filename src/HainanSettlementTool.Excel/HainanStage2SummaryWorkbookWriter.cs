using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using HainanSettlementTool.Core.Models;
using HainanSettlementTool.Core.Services;

namespace HainanSettlementTool.Excel
{
    internal static class HainanStage2SummaryWorkbookWriter
    {
        internal static void VerifyGeneratedSummary(
            HainanStage2Options options,
            IList<GroupSettlementTotal> totals,
            IList<HainanStage2SubjectGroup> subjectGroups,
            string summaryPath)
        {
            Dictionary<string, string> expectedPaymentPartyByKey;
            Dictionary<string, string> expectedPayeeByKey;
            HashSet<string> expectedMainKeys;
            using (var sourceStream = File.Open(options.SummaryTemplatePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var sourceWorkbook = new XLWorkbook(sourceStream))
            {
                var sourceMainName = ResolveSummarySheetName(sourceWorkbook, "main", true);
                var sourceMainMeta = ReadSummaryMeta(sourceWorkbook.Worksheet(sourceMainName));
                expectedPaymentPartyByKey = BuildPaymentPartyIndex(
                    options,
                    totals,
                    sourceWorkbook,
                    sourceMainMeta);
                expectedPayeeByKey = BuildCanonicalPayeeIndex(
                    totals,
                    sourceMainMeta,
                    ReadSummarySources(sourceWorkbook));
                expectedMainKeys = new HashSet<string>(sourceMainMeta.Select(SummaryKey));
                expectedMainKeys.UnionWith(totals.Select(total =>
                    HainanStage2ExcelUtil.SummaryKey(total.Entity, total.Kind)));
            }

            var expectedSubjectGroups = subjectGroups.ToDictionary(group =>
                HainanStage2ExcelUtil.SummaryKey(group.Entity, group.SettlementKind));
            var generatedTotals = totals
                .GroupBy(total => HainanStage2ExcelUtil.SummaryKey(total.Entity, total.Kind))
                .ToDictionary(group => group.Key, group => group.ToList());
            if (generatedTotals.Any(group => group.Value.Count != 1)
                || !new HashSet<string>(expectedSubjectGroups.Keys).SetEquals(generatedTotals.Keys))
            {
                throw new InvalidDataException("海南阶段二汇总表主体集合与台账不一致。");
            }

            using (var stream = File.Open(summaryPath, FileMode.Open, FileAccess.Read, FileShare.Read))
            using (var workbook = new XLWorkbook(stream))
            {
                var mainSheetName = ResolveSummarySheetName(workbook, "main", true);
                var mainSheet = workbook.Worksheet(mainSheetName);
                var monthColumns = FindSummaryMonthBlocks(mainSheet, options.Month);
                if (monthColumns.Count != 1)
                {
                    throw new InvalidDataException("海南阶段二主汇总表目标月块数量不是 1。");
                }

                var monthColumn = monthColumns[0];
                var mainRows = ReadSummaryMeta(mainSheet);
                var mainRowsByKey = mainRows
                    .GroupBy(SummaryKey)
                    .ToDictionary(group => group.Key, group => group.ToList());
                if (mainRowsByKey.Any(group => group.Value.Count != 1)
                    || !expectedMainKeys.SetEquals(mainRowsByKey.Keys))
                {
                    throw new InvalidDataException("海南阶段二主汇总表的主体集合或唯一性校验失败。");
                }

                var paymentSheets = new Dictionary<string, IXLWorksheet>();
                foreach (var paymentParty in HainanStage2PaymentParties.Supported)
                {
                    var role = paymentParty == HainanStage2PaymentParties.Qingneng ? "qingneng" : "qinghui";
                    var sheetName = ResolveSummarySheetName(workbook, role, false);
                    if (string.IsNullOrWhiteSpace(sheetName))
                    {
                        continue;
                    }

                    var sheet = workbook.Worksheet(sheetName);
                    if (FindSummaryMonthBlocks(sheet, options.Month).Count != 1)
                    {
                        throw new InvalidDataException("海南阶段二支付方汇总表目标月块数量不是 1：" + sheetName);
                    }

                    paymentSheets[paymentParty] = sheet;
                }

                foreach (var pair in mainRowsByKey)
                {
                    string expectedPaymentParty;
                    string expectedPayee;
                    if (!expectedPaymentPartyByKey.TryGetValue(pair.Key, out expectedPaymentParty)
                        || !expectedPayeeByKey.TryGetValue(pair.Key, out expectedPayee))
                    {
                        throw new InvalidDataException("海南阶段二汇总表缺少预检后的支付方或收款人结果：" + pair.Key);
                    }

                    var mainRow = pair.Value.Single();
                    if (mainRow.PaymentParty != expectedPaymentParty
                        || !Stage2OpaqueText.AreEquivalent(mainRow.Payee, expectedPayee))
                    {
                        throw new InvalidDataException("海南阶段二主汇总表支付方或完整收款人未按可靠来源生成：" + pair.Key);
                    }

                    VerifyPaymentPartyMembership(
                        paymentSheets,
                        pair.Key,
                        expectedPaymentParty,
                        mainSheet,
                        mainRow);
                }

                foreach (var total in totals)
                {
                    var key = HainanStage2ExcelUtil.SummaryKey(total.Entity, total.Kind);
                    var group = expectedSubjectGroups[key];
                    var row = mainRowsByKey[key].Single().Row;
                    var amountColumn = monthColumn + (total.Kind == "居间费" ? 1 : 0);
                    if (!AmountsEqual(ClosedXmlUtil.CellNumber(mainSheet.Cell(row, amountColumn)), total.ExpectedNet)
                        || !HainanStage2ExcelUtil.TaxRatesEqual(
                            ClosedXmlUtil.CellNumber(mainSheet.Cell(row, 9)),
                            group.TaxRate)
                        || !HainanStage2ExcelUtil.TaxRatesEqual(
                            ClosedXmlUtil.CellNumber(mainSheet.Cell(row, 10)),
                            0.13d)
                        || !string.Equals(
                            TextUtil.S(mainSheet.Cell(row, 8).FormulaA1),
                            "J" + row + "-I" + row,
                            StringComparison.OrdinalIgnoreCase)
                        || TextUtil.S(mainSheet.Cell(row, 11).GetFormattedString()) != TextUtil.S(group.Owner))
                    {
                        throw new InvalidDataException("海南阶段二主汇总表的金额、税率或负责人与台账不一致：" + total.Kind + " " + total.Entity);
                    }

                    var paymentParty = expectedPaymentPartyByKey[key];
                    IXLWorksheet paymentSheet;
                    if (!paymentSheets.TryGetValue(paymentParty, out paymentSheet))
                    {
                        throw new InvalidDataException("海南阶段二汇总表缺少本月主体对应的支付方工作表：" + paymentParty);
                    }

                    var paymentRow = ReadSummaryMeta(paymentSheet)
                        .Single(meta => SummaryKey(meta) == key)
                        .Row;
                    var paymentMonthColumn = FindSummaryMonthBlocks(paymentSheet, options.Month).Single();
                    var paymentAmountColumn = paymentMonthColumn + (total.Kind == "居间费" ? 1 : 0);
                    if (!AmountsEqual(ClosedXmlUtil.CellNumber(paymentSheet.Cell(paymentRow, paymentAmountColumn)), total.ExpectedNet)
                        || !HainanStage2ExcelUtil.TaxRatesEqual(
                            ClosedXmlUtil.CellNumber(paymentSheet.Cell(paymentRow, 9)),
                            group.TaxRate)
                        || !HainanStage2ExcelUtil.TaxRatesEqual(
                            ClosedXmlUtil.CellNumber(paymentSheet.Cell(paymentRow, 10)),
                            0.13d)
                        || !string.Equals(
                            TextUtil.S(paymentSheet.Cell(paymentRow, 8).FormulaA1),
                            "J" + paymentRow + "-I" + paymentRow,
                            StringComparison.OrdinalIgnoreCase)
                        || TextUtil.S(paymentSheet.Cell(paymentRow, 11).GetFormattedString()) != TextUtil.S(group.Owner))
                    {
                        throw new InvalidDataException("海南阶段二支付方汇总表未与主汇总和台账保持一致：" + total.Kind + " " + total.Entity);
                    }
                }
            }
        }

        internal static string BuildSummary(
            HainanStage2Options options,
            IList<GroupSettlementTotal> totals,
            IList<HainanStage2SubjectGroup> subjectGroups,
            IList<string> warnings)
        {
            var outputPath = PlanOutputPath(options);
            FileAccessGuard.RequireWritableWorkbook(outputPath, "输出汇总表");
            File.Copy(options.SummaryTemplatePath, outputPath, true);

            using (var workbook = new XLWorkbook(outputPath))
            {
                var mainSheetName = ResolveSummarySheetName(workbook, "main", true);
                var qingnengSheetName = ResolveSummarySheetName(workbook, "qingneng", false);
                var qinghuiSheetName = ResolveSummarySheetName(workbook, "qinghui", false);
                var mainMeta = ReadSummaryMeta(workbook.Worksheet(mainSheetName));
                var sourceMeta = ReadSummarySources(workbook);
                var partyByKey = BuildPaymentPartyIndex(options, totals, workbook, mainMeta);
                var payeeByKey = BuildCanonicalPayeeIndex(totals, mainMeta, sourceMeta);
                var taxRateByKey = subjectGroups.ToDictionary(
                    group => HainanStage2ExcelUtil.SummaryKey(group.Entity, group.SettlementKind),
                    group => group.TaxRate);
                var knownMainKeys = new HashSet<string>(mainMeta.Select(SummaryKey));
                var newSubjectKeys = new HashSet<string>(totals
                    .Select(total => HainanStage2ExcelUtil.SummaryKey(total.Entity, total.Kind))
                    .Where(key => !knownMainKeys.Contains(key)));

                var mainSheet = workbook.Worksheet(mainSheetName);
                WriteSummarySheet(mainSheet, totals, options.Month, null, warnings, partyByKey, payeeByKey, taxRateByKey, newSubjectKeys, null);

                if (!string.IsNullOrWhiteSpace(qingnengSheetName))
                {
                    var qnTotals = totals.Where(total => PartyForSummaryTotal(total, partyByKey) == HainanStage2PaymentParties.Qingneng).ToList();
                    var allowed = new HashSet<string>(partyByKey.Where(item => item.Value == HainanStage2PaymentParties.Qingneng).Select(item => item.Key));
                    foreach (var total in qnTotals)
                    {
                        allowed.Add(HainanStage2ExcelUtil.SummaryKey(total.Entity, total.Kind));
                    }
                    WriteSummarySheet(workbook.Worksheet(qingnengSheetName), qnTotals, options.Month, allowed, warnings, partyByKey, payeeByKey, taxRateByKey, newSubjectKeys, mainSheet);
                }

                if (!string.IsNullOrWhiteSpace(qinghuiSheetName))
                {
                    var qhTotals = totals.Where(total => PartyForSummaryTotal(total, partyByKey) == HainanStage2PaymentParties.Qinghui).ToList();
                    var allowed = new HashSet<string>(partyByKey.Where(item => item.Value == HainanStage2PaymentParties.Qinghui).Select(item => item.Key));
                    foreach (var total in qhTotals)
                    {
                        allowed.Add(HainanStage2ExcelUtil.SummaryKey(total.Entity, total.Kind));
                    }
                    WriteSummarySheet(workbook.Worksheet(qinghuiSheetName), qhTotals, options.Month, allowed, warnings, partyByKey, payeeByKey, taxRateByKey, newSubjectKeys, mainSheet);
                }

                HainanStage2ExcelUtil.SaveWorkbook(workbook, outputPath);
            }

            return outputPath;
        }

        internal static string PlanOutputPath(HainanStage2Options options)
        {
            var outputName = string.IsNullOrWhiteSpace(options.OutputSummaryName)
                ? "【2026年海南省代理费汇总表-" + options.Month + "月自动化】.xlsx"
                : options.OutputSummaryName;
            var outputRoot = Path.GetFullPath(options.OutputDirectory);
            var outputPath = Path.GetFullPath(Path.Combine(outputRoot, outputName));
            var outputRootPrefix = outputRoot
                .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
                + Path.DirectorySeparatorChar;
            if (!outputPath.StartsWith(outputRootPrefix, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(
                    "海南阶段二汇总表输出路径必须严格位于本次输出根目录内。");
            }

            return outputPath;
        }

        private static void WriteSummarySheet(
            IXLWorksheet worksheet,
            IList<GroupSettlementTotal> totals,
            int month,
            ISet<string> allowedKeys,
            IList<string> warnings,
            IDictionary<string, string> paymentPartyByKey,
            IDictionary<string, string> payeeByKey,
            IDictionary<string, double> taxRateByKey,
            ISet<string> newSubjectKeys,
            IXLWorksheet canonicalMainSheet)
        {
            var startRow = 4;
            DeleteSummaryRowsNotAllowed(worksheet, startRow, allowedKeys);

            var monthColumn = InsertMonthBlock(worksheet, month);
            var cumulativeColumn = SummaryColumn(worksheet, "累计代理费总计");
            var totalByKey = totals.ToDictionary(total => HainanStage2ExcelUtil.SummaryKey(total.Entity, total.Kind), total => total);
            var newTotals = totals
                .Where(total => newSubjectKeys.Contains(HainanStage2ExcelUtil.SummaryKey(total.Entity, total.Kind)))
                .ToList();
            if (canonicalMainSheet == null)
            {
                InsertNewSummaryRows(worksheet, startRow, newTotals, warnings, paymentPartyByKey);
            }
            else
            {
                InsertMissingCanonicalRows(worksheet, canonicalMainSheet, startRow, allowedKeys);
                SynchronizeCanonicalBusinessFields(worksheet, canonicalMainSheet);
            }

            var dataRows = ReadSummaryMeta(worksheet).OrderBy(item => item.Row).ToList();
            for (var index = 0; index < dataRows.Count; index++)
            {
                var row = dataRows[index].Row;
                var info = dataRows[index];
                GroupSettlementTotal total;
                totalByKey.TryGetValue(HainanStage2ExcelUtil.SummaryKey(info.Entity, info.Kind), out total);
                worksheet.Cell(row, 1).Value = index + 1;
                WriteSummaryValues(
                    worksheet,
                    row,
                    monthColumn,
                    cumulativeColumn,
                    total,
                    info.Entity,
                    info.Kind,
                    month,
                    newSubjectKeys.Contains(HainanStage2ExcelUtil.SummaryKey(info.Entity, info.Kind)),
                    paymentPartyByKey,
                    payeeByKey,
                    taxRateByKey);
            }

            var totalRow = FindSummaryTotalRow(worksheet, startRow);
            worksheet.Cell(totalRow, 1).Value = "合计";
            WriteSummaryTotalRow(worksheet, totalRow, cumulativeColumn);
            ApplySummaryDateFormats(worksheet, startRow, totalRow, cumulativeColumn);
            UpdateSummarySignatureDate(worksheet, month);
            ApplySummaryHeaderMerges(worksheet, monthColumn, cumulativeColumn + 9);
        }

        private static void WriteSummaryValues(
            IXLWorksheet worksheet,
            int row,
            int monthColumn,
            int cumulativeColumn,
            GroupSettlementTotal total,
            string entity,
            string kind,
            int month,
            bool isNewRow,
            IDictionary<string, string> paymentPartyByKey,
            IDictionary<string, string> payeeByKey,
            IDictionary<string, double> taxRateByKey)
        {
            var proxyValue = total != null && total.Kind == "代理费" ? total.ExpectedNet : (double?)null;
            var interValue = total != null && total.Kind == "居间费" ? total.ExpectedNet : (double?)null;
            SetNullableNumber(worksheet.Cell(row, monthColumn), proxyValue);
            SetNullableNumber(worksheet.Cell(row, monthColumn + 1), interValue);
            worksheet.Cell(row, monthColumn + 2).Clear(XLClearOptions.Contents);

            var fee = (proxyValue ?? 0) + (interValue ?? 0);
            var loanTotalCell = worksheet.Cell(row, cumulativeColumn + 1);
            var hasLoan = CellHasContent(loanTotalCell);
            var loanTotal = hasLoan ? ClosedXmlUtil.CellNumber(loanTotalCell) : 0;
            var previousDeducted = hasLoan ? ClosedXmlUtil.CellNumber(worksheet.Cell(row, cumulativeColumn + 2)) : 0;
            var monthlyDeduction = ParseMonthlyDeduction(TextUtil.S(worksheet.Cell(row, cumulativeColumn + 5).GetFormattedString()));
            var remaining = Math.Max(loanTotal - previousDeducted, 0);
            var deduction = 0d;
            if (remaining > 0 && fee > 0)
            {
                deduction = Math.Round(Math.Min(Math.Min(fee, remaining), monthlyDeduction == 0 ? remaining : monthlyDeduction), 4);
            }

            SetNullableNumber(worksheet.Cell(row, monthColumn + 3), deduction == 0 ? (double?)null : deduction);
            worksheet.Cell(row, monthColumn + 4).FormulaA1 = SumFormula(row, monthColumn, monthColumn + 2);
            worksheet.Cell(row, monthColumn + 5).FormulaA1 = ClosedXmlUtil.ColumnLetter(monthColumn + 4) + row + "-" + ClosedXmlUtil.ColumnLetter(monthColumn + 3) + row;
            worksheet.Cell(row, cumulativeColumn).FormulaA1 = SumEverySix(row, 16, cumulativeColumn - 1);
            if (hasLoan)
            {
                worksheet.Cell(row, cumulativeColumn + 2).FormulaA1 = SumEverySix(row, 15, cumulativeColumn - 1);
                worksheet.Cell(row, cumulativeColumn + 3).FormulaA1 = ClosedXmlUtil.ColumnLetter(cumulativeColumn + 1) + row + "-" + ClosedXmlUtil.ColumnLetter(cumulativeColumn + 2) + row;
            }
            else
            {
                worksheet.Cell(row, cumulativeColumn + 2).Clear(XLClearOptions.Contents);
                worksheet.Cell(row, cumulativeColumn + 3).Clear(XLClearOptions.Contents);
            }

            if (isNewRow && !CellHasContent(worksheet.Cell(row, cumulativeColumn + 7)))
            {
                worksheet.Cell(row, cumulativeColumn + 7).Value = new DateTime(HainanStage2ExcelUtil.Year, month, 1);
            }

            if (total != null)
            {
                double taxRate;
                if (!taxRateByKey.TryGetValue(HainanStage2ExcelUtil.SummaryKey(entity, kind), out taxRate))
                {
                    throw new InvalidOperationException("海南阶段二汇总主体缺少唯一扣税率：" + kind + " " + entity + "。");
                }

                worksheet.Cell(row, 8).FormulaA1 = "J" + row + "-I" + row;
                worksheet.Cell(row, 9).Value = taxRate;
                worksheet.Cell(row, 10).Value = 0.13;
                worksheet.Cell(row, 11).Value = total.Owner;
            }

            worksheet.Cell(row, 6).Value = PayeeFromIndex(payeeByKey, entity, kind);
            worksheet.Cell(row, cumulativeColumn + 8).Value = PaymentPartyFromIndex(paymentPartyByKey, entity, kind);
        }

        private static int InsertMonthBlock(IXLWorksheet worksheet, int month)
        {
            var existing = FindSummaryMonthBlocks(worksheet, month);
            var matchingHeaders = FindSummaryMonthHeaderColumns(worksheet, month);
            if (matchingHeaders.Count > 1 || matchingHeaders.Count != existing.Count)
            {
                throw new InvalidDataException(worksheet.Name + "中的" + month + "月标题重复或结构不完整，无法安全重写。");
            }

            if (existing.Count == 1)
            {
                return existing[0];
            }

            var cumulativeColumn = SummaryColumn(worksheet, "累计代理费总计");
            var insertAt = cumulativeColumn;
            worksheet.Column(insertAt).InsertColumnsBefore(6);
            for (var offset = 0; offset < 6; offset++)
            {
                var sourceColumn = insertAt - 6 + offset;
                var targetColumn = insertAt + offset;
                worksheet.Column(sourceColumn).CopyTo(worksheet.Column(targetColumn));
                worksheet.Column(targetColumn).Unhide();
            }

            for (var column = insertAt - 6; column < insertAt; column++)
            {
                worksheet.Column(column).Hide();
            }

            for (var column = insertAt; column < insertAt + 6; column++)
            {
                worksheet.Column(column).Unhide();
            }

            worksheet.Cell(2, insertAt).Value = HainanStage2ExcelUtil.Year + "年" + month + "月";
            var labels = new[] { "代理费", "居间费", "退补电费", "当月抵扣", "费用合计" };
            for (var index = 0; index < labels.Length; index++)
            {
                worksheet.Cell(3, insertAt + index).Value = labels[index];
            }
            worksheet.Cell(2, insertAt + 5).Value = "当月实际支付";
            worksheet.Cell(3, insertAt + 5).Clear(XLClearOptions.Contents);
            return insertAt;
        }

        private static void DeleteSummaryRowsNotAllowed(IXLWorksheet worksheet, int startRow, ISet<string> allowedKeys)
        {
            if (allowedKeys == null)
            {
                return;
            }

            var totalRow = FindSummaryTotalRow(worksheet, startRow);
            for (var row = totalRow - 1; row >= startRow; row--)
            {
                var entity = TextUtil.S(worksheet.Cell(row, 2).GetFormattedString());
                var kind = TextUtil.S(worksheet.Cell(row, 3).GetFormattedString());
                if (string.IsNullOrWhiteSpace(entity) || !allowedKeys.Contains(HainanStage2ExcelUtil.SummaryKey(entity, kind)))
                {
                    worksheet.Row(row).Delete();
                }
            }
        }

        private static void InsertNewSummaryRows(
            IXLWorksheet worksheet,
            int startRow,
            IList<GroupSettlementTotal> newTotals,
            IList<string> warnings,
            IDictionary<string, string> paymentPartyByKey)
        {
            if (newTotals.Count == 0)
            {
                return;
            }

            var totalRow = FindSummaryTotalRow(worksheet, startRow);
            var templateRow = Math.Max(startRow, totalRow - 1);
            foreach (var total in newTotals)
            {
                worksheet.Row(totalRow).InsertRowsAbove(1);
                worksheet.Row(templateRow).CopyTo(worksheet.Row(totalRow));
                worksheet.Row(totalRow).Clear(XLClearOptions.Contents);
                worksheet.Cell(totalRow, 1).Value = totalRow - 3;
                worksheet.Cell(totalRow, 2).Value = total.Entity;
                worksheet.Cell(totalRow, 3).Value = total.Kind;
                worksheet.Cell(totalRow, 4).Value = "否";
                worksheet.Cell(totalRow, 6).Value = total.Entity;
                worksheet.Cell(totalRow, 7).Value = "平台";
                worksheet.Cell(totalRow, 8).FormulaA1 = "J" + totalRow + "-I" + totalRow;
                worksheet.Cell(totalRow, 10).Value = 0.13;
                worksheet.Cell(totalRow, 11).Value = total.Owner;
                warnings.Add("新增汇总主体：" + total.Kind + " " + total.Entity + "（负责人：" + total.Owner + "；支付方：" + PaymentPartyFromIndex(paymentPartyByKey, total.Entity, total.Kind) + "）");
                totalRow++;
            }
        }

        private static void InsertMissingCanonicalRows(
            IXLWorksheet worksheet,
            IXLWorksheet canonicalMainSheet,
            int startRow,
            ISet<string> allowedKeys)
        {
            if (allowedKeys == null)
            {
                return;
            }

            var knownKeys = new HashSet<string>(ReadSummaryMeta(worksheet)
                .Select(item => HainanStage2ExcelUtil.SummaryKey(item.Entity, item.Kind)));
            var missing = ReadSummaryMeta(canonicalMainSheet)
                .Where(item =>
                {
                    var key = HainanStage2ExcelUtil.SummaryKey(item.Entity, item.Kind);
                    return allowedKeys.Contains(key) && !knownKeys.Contains(key);
                })
                .OrderBy(item => item.Row)
                .ToList();
            var totalRow = FindSummaryTotalRow(worksheet, startRow);
            foreach (var source in missing)
            {
                worksheet.Row(totalRow).InsertRowsAbove(1);
                canonicalMainSheet.Row(source.Row).CopyTo(worksheet.Row(totalRow));
                knownKeys.Add(HainanStage2ExcelUtil.SummaryKey(source.Entity, source.Kind));
                totalRow++;
            }
        }

        private static void SynchronizeCanonicalBusinessFields(
            IXLWorksheet worksheet,
            IXLWorksheet canonicalMainSheet)
        {
            var canonicalRows = ReadSummaryMeta(canonicalMainSheet)
                .GroupBy(SummaryKey)
                .ToDictionary(group => group.Key, group => group.Single());
            var paymentRows = ReadSummaryMeta(worksheet);
            var canonicalCumulativeColumn = SummaryColumn(canonicalMainSheet, "累计代理费总计");
            var paymentCumulativeColumn = SummaryColumn(worksheet, "累计代理费总计");
            foreach (var paymentRow in paymentRows)
            {
                HainanStage2SummaryMetaRow canonical;
                if (!canonicalRows.TryGetValue(SummaryKey(paymentRow), out canonical))
                {
                    continue;
                }

                for (var column = 2; column <= 7; column++)
                {
                    CopyBusinessValue(
                        canonicalMainSheet.Cell(canonical.Row, column),
                        worksheet.Cell(paymentRow.Row, column));
                }

                var canonicalTaxFormula = TextUtil.S(canonicalMainSheet.Cell(canonical.Row, 8).FormulaA1);
                if (string.IsNullOrWhiteSpace(canonicalTaxFormula))
                {
                    CopyBusinessValue(
                        canonicalMainSheet.Cell(canonical.Row, 8),
                        worksheet.Cell(paymentRow.Row, 8));
                }
                else
                {
                    worksheet.Cell(paymentRow.Row, 8).FormulaA1 = "J" + paymentRow.Row + "-I" + paymentRow.Row;
                }

                for (var column = 9; column <= 11; column++)
                {
                    CopyBusinessValue(
                        canonicalMainSheet.Cell(canonical.Row, column),
                        worksheet.Cell(paymentRow.Row, column));
                }

                foreach (var offset in new[] { 1, 4, 5, 6, 7, 9 })
                {
                    CopyBusinessValue(
                        canonicalMainSheet.Cell(canonical.Row, canonicalCumulativeColumn + offset),
                        worksheet.Cell(paymentRow.Row, paymentCumulativeColumn + offset));
                }
            }
        }

        private static void CopyBusinessValue(IXLCell source, IXLCell target)
        {
            target.Clear(XLClearOptions.Contents);
            if (source.HasFormula)
            {
                target.FormulaR1C1 = source.FormulaR1C1;
            }
            else if (!source.IsEmpty())
            {
                target.Value = source.Value;
            }
        }

        private static void WriteSummaryTotalRow(IXLWorksheet worksheet, int totalRow, int cumulativeColumn)
        {
            for (var column = 12; column <= cumulativeColumn + 3; column++)
            {
                var letter = ClosedXmlUtil.ColumnLetter(column);
                worksheet.Cell(totalRow, column).FormulaA1 = "SUM(" + letter + "4:" + letter + (totalRow - 1) + ")";
            }

            for (var column = cumulativeColumn + 4; column <= Math.Min(cumulativeColumn + 9, worksheet.LastColumnUsed()?.ColumnNumber() ?? cumulativeColumn + 9); column++)
            {
                worksheet.Cell(totalRow, column).Clear(XLClearOptions.Contents);
            }
        }

        private static void ApplySummaryDateFormats(IXLWorksheet worksheet, int startRow, int totalRow, int cumulativeColumn)
        {
            foreach (var column in new[] { cumulativeColumn + 4, cumulativeColumn + 6, cumulativeColumn + 7 })
            {
                for (var row = startRow; row < totalRow; row++)
                {
                    if (CellHasContent(worksheet.Cell(row, column)))
                    {
                        worksheet.Cell(row, column).Style.DateFormat.Format = "yyyy年m月";
                    }
                }
            }
        }

        private static void UpdateSummarySignatureDate(IXLWorksheet worksheet, int month)
        {
            var date = new DateTime(HainanStage2ExcelUtil.Year, month, 8).AddMonths(2);
            var text = "日期：" + date.Year + "年" + date.Month.ToString("00") + "月" + date.Day.ToString("00") + "日";
            IXLCell target = null;
            var maxColumn = 0;
            foreach (var cell in worksheet.CellsUsed())
            {
                if (!TextUtil.S(cell.GetFormattedString()).Contains("日期："))
                {
                    continue;
                }

                if (cell.Address.ColumnNumber > maxColumn)
                {
                    target = cell;
                    maxColumn = cell.Address.ColumnNumber;
                }
            }

            if (target != null)
            {
                target.Value = text;
            }
        }

        private static void ApplySummaryHeaderMerges(IXLWorksheet worksheet, int monthColumn, int lastRelevantColumn)
        {
            UnmergeIntersecting(worksheet, 1, 1, 1, lastRelevantColumn);
            worksheet.Range(1, 1, 1, lastRelevantColumn).Merge();

            UnmergeIntersecting(worksheet, 2, monthColumn, 2, monthColumn + 4);
            worksheet.Range(2, monthColumn, 2, monthColumn + 4).Merge();

            UnmergeIntersecting(worksheet, 2, monthColumn + 5, 3, monthColumn + 5);
            worksheet.Range(2, monthColumn + 5, 3, monthColumn + 5).Merge();
            worksheet.Cell(2, monthColumn + 5).Value = "当月实际支付";
        }

        private static void UnmergeIntersecting(IXLWorksheet worksheet, int firstRow, int firstColumn, int lastRow, int lastColumn)
        {
            var ranges = worksheet.MergedRanges
                .Where(range => RangeIntersects(range, firstRow, firstColumn, lastRow, lastColumn))
                .ToList();
            foreach (var range in ranges)
            {
                range.Unmerge();
            }
        }

        private static bool RangeIntersects(IXLRange range, int firstRow, int firstColumn, int lastRow, int lastColumn)
        {
            var address = range.RangeAddress;
            return address.FirstAddress.RowNumber <= lastRow
                && address.LastAddress.RowNumber >= firstRow
                && address.FirstAddress.ColumnNumber <= lastColumn
                && address.LastAddress.ColumnNumber >= firstColumn;
        }

        internal static IList<HainanStage2SummaryMetaRow> ReadSummaryMeta(IXLWorksheet worksheet)
        {
            var invalidKinds = FindInvalidSummaryKindRows(worksheet);
            if (invalidKinds.Count > 0)
            {
                var invalid = invalidKinds[0];
                throw new InvalidDataException(
                    worksheet.Name + "第" + invalid.Row + "行主体“" + invalid.Entity
                    + "”的费用类型为“" + (string.IsNullOrWhiteSpace(invalid.Kind) ? "空白" : invalid.Kind)
                    + "”，只允许代理费或居间费。");
            }

            var cumulativeColumn = SummaryColumn(worksheet, "累计代理费总计");
            var totalRow = FindSummaryTotalRow(worksheet, 4);
            var rows = new List<HainanStage2SummaryMetaRow>();
            for (var row = 4; row < totalRow; row++)
            {
                var entity = TextUtil.S(worksheet.Cell(row, 2).GetFormattedString());
                var kind = TextUtil.S(worksheet.Cell(row, 3).GetFormattedString());
                if (string.IsNullOrWhiteSpace(entity) || entity.Contains("审核"))
                {
                    continue;
                }

                rows.Add(new HainanStage2SummaryMetaRow
                {
                    Row = row,
                    SheetName = worksheet.Name,
                    Entity = entity,
                    Kind = kind,
                    Payee = worksheet.Cell(row, 6).GetFormattedString(),
                    PaymentParty = TextUtil.S(worksheet.Cell(row, cumulativeColumn + 8).GetFormattedString())
                });
            }

            return rows;
        }

        internal static IList<HainanStage2SummaryMetaRow> FindInvalidSummaryKindRows(
            IXLWorksheet worksheet)
        {
            var totalRow = FindSummaryTotalRow(worksheet, 4);
            var rows = new List<HainanStage2SummaryMetaRow>();
            for (var row = 4; row < totalRow; row++)
            {
                var entity = TextUtil.S(worksheet.Cell(row, 2).GetFormattedString());
                if (string.IsNullOrWhiteSpace(entity) || entity.Contains("审核"))
                {
                    continue;
                }

                var kind = TextUtil.S(worksheet.Cell(row, 3).GetFormattedString());
                if (kind == "代理费" || kind == "居间费")
                {
                    continue;
                }

                rows.Add(new HainanStage2SummaryMetaRow
                {
                    Row = row,
                    SheetName = worksheet.Name,
                    Entity = entity,
                    Kind = kind
                });
            }

            return rows;
        }

        internal static IList<HainanStage2SummaryMetaRow> ReadSummarySources(XLWorkbook workbook)
        {
            var sheetNames = new[]
            {
                ResolveSummarySheetName(workbook, "main", true),
                ResolveSummarySheetName(workbook, "qingneng", false),
                ResolveSummarySheetName(workbook, "qinghui", false)
            };
            var rows = new List<HainanStage2SummaryMetaRow>();
            foreach (var sheetName in sheetNames
                .Where(name => !string.IsNullOrWhiteSpace(name))
                .Distinct(StringComparer.Ordinal))
            {
                rows.AddRange(ReadSummaryMeta(workbook.Worksheet(sheetName)));
            }

            return rows;
        }

        internal static IList<int> FindSummaryMonthBlocks(IXLWorksheet worksheet, int month)
        {
            return FindSummaryMonthHeaderColumns(worksheet, month)
                .Where(column => TextUtil.S(worksheet.Cell(3, column).GetFormattedString()) == "代理费")
                .ToList();
        }

        internal static IList<int> FindSummaryMonthHeaderColumns(IXLWorksheet worksheet, int month)
        {
            var result = new List<int>();
            var lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? 1;
            for (var column = 1; column <= lastColumn; column++)
            {
                var header = Regex.Replace(TextUtil.S(worksheet.Cell(2, column).GetFormattedString()), "\\s+", string.Empty);
                if (header == HainanStage2ExcelUtil.Year + "年" + month + "月"
                    || header == HainanStage2ExcelUtil.Year + "年" + month.ToString("00") + "月")
                {
                    result.Add(column);
                }
            }

            return result;
        }

        private static bool CellHasContent(IXLCell cell)
        {
            if (!string.IsNullOrWhiteSpace(cell.FormulaA1))
            {
                return true;
            }

            return !string.IsNullOrWhiteSpace(cell.GetFormattedString());
        }

        internal static string ResolveSummarySheetName(XLWorkbook workbook, string role, bool required)
        {
            var names = workbook.Worksheets.Select(sheet => sheet.Name).ToList();
            var candidates = FindSummarySheetCandidates(workbook, role);
            if (candidates.Count == 1)
            {
                return candidates[0];
            }

            if (candidates.Count > 1)
            {
                throw new InvalidDataException(
                    "选择的汇总表模板中" + role + "角色匹配到多个工作表："
                    + string.Join("、", candidates) + "。请只保留一个权威工作表。");
            }

            if (required)
            {
                throw new InvalidOperationException("选择的汇总表模板缺少" + role + "汇总表。当前工作表：" + string.Join("、", names) + "。请在“上月/修正版汇总表”选择代理费汇总表文件。");
            }

            return null;
        }

        internal static IList<string> FindSummarySheetCandidates(XLWorkbook workbook, string role)
        {
            var names = workbook.Worksheets.Select(sheet => sheet.Name).ToList();
            Func<string, string> clean = name => Regex.Replace(TextUtil.S(name), "\\s+", string.Empty);
            if (role == "main")
            {
                return names
                    .Where(name => clean(name).Contains("汇总表")
                        && !clean(name).Contains("清能")
                        && !clean(name).Contains("清辉")
                        && SummarySheetHasMarker(workbook.Worksheet(name)))
                    .ToList();
            }

            if (role == "qingneng" || role == "qinghui")
            {
                var marker = role == "qingneng" ? "清能" : "清辉";
                return names
                    .Where(name => clean(name).Contains(marker)
                        && clean(name).Contains("汇总")
                        && SummarySheetHasMarker(workbook.Worksheet(name)))
                    .ToList();
            }

            throw new ArgumentException("未知的海南阶段二汇总工作表角色：" + role, nameof(role));
        }

        private static bool SummarySheetHasMarker(IXLWorksheet worksheet)
        {
            try
            {
                SummaryColumn(worksheet, "累计代理费总计");
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static int SummaryColumn(IXLWorksheet worksheet, string header)
        {
            var lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? 1;
            for (var column = 1; column <= lastColumn; column++)
            {
                if (TextUtil.S(worksheet.Cell(2, column).GetFormattedString()) == header)
                {
                    return column;
                }
            }

            throw new InvalidOperationException(worksheet.Name + " 未找到列：" + header);
        }

        private static int FindSummaryTotalRow(IXLWorksheet worksheet, int startRow)
        {
            var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? startRow;
            for (var row = startRow; row <= lastRow; row++)
            {
                if (TextUtil.S(worksheet.Cell(row, 1).GetFormattedString()) == "合计")
                {
                    return row;
                }
            }

            for (var row = startRow; row <= lastRow; row++)
            {
                for (var column = 12; column <= Math.Min(worksheet.LastColumnUsed()?.ColumnNumber() ?? 12, 80); column++)
                {
                    var formula = TextUtil.S(worksheet.Cell(row, column).FormulaA1);
                    if (formula.StartsWith("SUM(", StringComparison.OrdinalIgnoreCase))
                    {
                        return row;
                    }
                }
            }

            throw new InvalidOperationException(worksheet.Name + " 未找到合计行。");
        }

        private static Dictionary<string, string> BuildPaymentPartyIndex(
            HainanStage2Options options,
            IList<GroupSettlementTotal> totals,
            XLWorkbook workbook,
            IList<HainanStage2SummaryMetaRow> mainMeta)
        {
            var sources = ReadSummarySources(workbook);
            var qingnengSheetName = ResolveSummarySheetName(workbook, "qingneng", false);
            var qinghuiSheetName = ResolveSummarySheetName(workbook, "qinghui", false);
            var partyByKey = new Dictionary<string, string>();
            foreach (var item in mainMeta)
            {
                var key = HainanStage2ExcelUtil.SummaryKey(item.Entity, item.Kind);
                if (partyByKey.ContainsKey(key))
                {
                    throw new InvalidOperationException("海南阶段二汇总表存在重复主体：" + item.Kind + " " + item.Entity + "。");
                }

                string paymentParty;
                if (!TryResolvePaymentParty(
                    options,
                    item.Entity,
                    item.Kind,
                    sources,
                    qingnengSheetName,
                    qinghuiSheetName,
                    out paymentParty))
                {
                    throw new InvalidOperationException("海南阶段二汇总主体支付方未选择或存在冲突：" + item.Kind + " " + item.Entity + "。请返回预检处理。");
                }

                partyByKey[key] = paymentParty;
            }

            foreach (var total in totals)
            {
                var key = HainanStage2ExcelUtil.SummaryKey(total.Entity, total.Kind);
                if (partyByKey.ContainsKey(key))
                {
                    continue;
                }

                string paymentParty;
                if (!TryResolvePaymentParty(
                    options,
                    total.Entity,
                    total.Kind,
                    sources,
                    qingnengSheetName,
                    qinghuiSheetName,
                    out paymentParty))
                {
                    throw new InvalidOperationException("海南阶段二新增汇总主体支付方未选择：" + total.Kind + " " + total.Entity + "。请在预检窗口选择清能或清辉后再生成。");
                }

                partyByKey[key] = paymentParty;
            }

            return partyByKey;
        }

        private static Dictionary<string, string> BuildCanonicalPayeeIndex(
            IList<GroupSettlementTotal> totals,
            IList<HainanStage2SummaryMetaRow> mainMeta,
            IList<HainanStage2SummaryMetaRow> sources)
        {
            var result = new Dictionary<string, string>();
            foreach (var item in mainMeta)
            {
                var key = HainanStage2ExcelUtil.SummaryKey(item.Entity, item.Kind);
                if (result.ContainsKey(key))
                {
                    throw new InvalidOperationException("海南阶段二汇总表存在重复主体：" + item.Kind + " " + item.Entity + "。");
                }

                if (!string.IsNullOrWhiteSpace(Stage2OpaqueText.NormalizeForComparison(item.Payee)))
                {
                    result[key] = item.Payee;
                    continue;
                }

                var candidates = sources
                    .Where(source => HainanStage2ExcelUtil.SummaryKey(source.Entity, source.Kind) == key)
                    .Where(source => !string.IsNullOrWhiteSpace(Stage2OpaqueText.NormalizeForComparison(source.Payee)))
                    .GroupBy(source => Stage2OpaqueText.NormalizeForComparison(source.Payee))
                    .Select(group => group.First().Payee)
                    .ToList();
                result[key] = candidates.Count == 1 ? candidates[0] : string.Empty;
            }

            foreach (var total in totals)
            {
                var key = HainanStage2ExcelUtil.SummaryKey(total.Entity, total.Kind);
                if (!result.ContainsKey(key))
                {
                    result[key] = total.Entity;
                }
            }

            return result;
        }

        private static bool TryResolvePaymentParty(
            HainanStage2Options options,
            string entity,
            string kind,
            IEnumerable<HainanStage2SummaryMetaRow> sources,
            string qingnengSheetName,
            string qinghuiSheetName,
            out string paymentParty)
        {
            if (HainanStage2ExcelUtil.TryGetPaymentPartyOverride(entity, kind, options.Month, out paymentParty))
            {
                return true;
            }

            var key = HainanStage2ExcelUtil.SummaryKey(entity, kind);
            var matchingSources = sources
                .Where(item => HainanStage2ExcelUtil.SummaryKey(item.Entity, item.Kind) == key)
                .ToList();
            var inQingneng = !string.IsNullOrWhiteSpace(qingnengSheetName)
                && matchingSources.Any(item => item.SheetName == qingnengSheetName);
            var inQinghui = !string.IsNullOrWhiteSpace(qinghuiSheetName)
                && matchingSources.Any(item => item.SheetName == qinghuiSheetName);
            if (inQingneng && inQinghui)
            {
                paymentParty = null;
                return false;
            }

            if (inQingneng || inQinghui)
            {
                var membershipParty = inQingneng
                    ? HainanStage2PaymentParties.Qingneng
                    : HainanStage2PaymentParties.Qinghui;
                var hasConflictingField = matchingSources
                    .Select(item => Stage2OpaqueText.NormalizeForComparison(item.PaymentParty))
                    .Any(value => !string.IsNullOrWhiteSpace(value) && value != membershipParty);
                if (hasConflictingField)
                {
                    paymentParty = null;
                    return false;
                }

                paymentParty = membershipParty;
                return true;
            }

            if (HainanStage2ExcelUtil.TryGetPaymentPartyDecision(options, entity, kind, out paymentParty))
            {
                return HainanStage2PaymentParties.Supported.Contains(paymentParty);
            }

            var values = matchingSources
                .Select(item => Stage2OpaqueText.NormalizeForComparison(item.PaymentParty))
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Distinct(StringComparer.Ordinal)
                .ToList();
            if (values.Count == 1 && HainanStage2PaymentParties.Supported.Contains(values[0]))
            {
                paymentParty = values[0];
                return true;
            }

            paymentParty = null;
            return false;
        }

        private static string PartyForSummaryTotal(GroupSettlementTotal total, IDictionary<string, string> partyByKey)
        {
            string party;
            if (partyByKey.TryGetValue(HainanStage2ExcelUtil.SummaryKey(total.Entity, total.Kind), out party))
            {
                return party;
            }

            throw new InvalidOperationException("海南阶段二新增汇总主体支付方未选择：" + total.Kind + " " + total.Entity + "。请在预检窗口选择清能或清辉后再生成。");
        }

        private static string PaymentPartyFromIndex(IDictionary<string, string> paymentPartyByKey, string entity, string kind)
        {
            string paymentParty;
            if (paymentPartyByKey != null && paymentPartyByKey.TryGetValue(HainanStage2ExcelUtil.SummaryKey(entity, kind), out paymentParty))
            {
                return paymentParty;
            }

            throw new InvalidOperationException("海南阶段二新增汇总主体支付方未选择：" + kind + " " + entity + "。请在预检窗口选择清能或清辉后再生成。");
        }

        private static string PayeeFromIndex(IDictionary<string, string> payeeByKey, string entity, string kind)
        {
            string payee;
            if (payeeByKey != null && payeeByKey.TryGetValue(HainanStage2ExcelUtil.SummaryKey(entity, kind), out payee))
            {
                return payee ?? string.Empty;
            }

            throw new InvalidOperationException("海南阶段二汇总主体缺少可靠收款人来源：" + kind + " " + entity + "。");
        }

        private static void VerifyPaymentPartyMembership(
            IDictionary<string, IXLWorksheet> paymentSheets,
            string key,
            string expectedPaymentParty,
            IXLWorksheet mainSheet,
            HainanStage2SummaryMetaRow mainRow)
        {
            foreach (var paymentParty in HainanStage2PaymentParties.Supported)
            {
                IXLWorksheet sheet;
                if (!paymentSheets.TryGetValue(paymentParty, out sheet))
                {
                    continue;
                }

                var matches = ReadSummaryMeta(sheet)
                    .Where(row => SummaryKey(row) == key)
                    .ToList();
                var expectedCount = paymentParty == expectedPaymentParty ? 1 : 0;
                if (matches.Count != expectedCount)
                {
                    throw new InvalidDataException("海南阶段二支付方工作表的主体归属不互斥：" + key);
                }

                if (matches.Count == 1)
                {
                    var paymentRow = matches[0];
                    if (!Stage2OpaqueText.AreEquivalent(paymentRow.Payee, mainRow.Payee)
                        || paymentRow.PaymentParty != mainRow.PaymentParty)
                    {
                        throw new InvalidDataException("海南阶段二支付方工作表的完整收款人或支付方未与主汇总保持一致：" + key);
                    }

                    VerifyCanonicalBusinessFields(mainSheet, mainRow.Row, sheet, paymentRow.Row, key);
                }
            }
        }

        private static void VerifyCanonicalBusinessFields(
            IXLWorksheet mainSheet,
            int mainRow,
            IXLWorksheet paymentSheet,
            int paymentRow,
            string key)
        {
            foreach (var column in new[] { 2, 3, 4, 5, 7 })
            {
                if (!BusinessValuesEqual(mainSheet.Cell(mainRow, column), paymentSheet.Cell(paymentRow, column)))
                {
                    throw new InvalidDataException("海南阶段二支付方工作表长期业务字段未与主汇总同步：" + key);
                }
            }

            var mainTaxFormula = TextUtil.S(mainSheet.Cell(mainRow, 8).FormulaA1);
            if (string.IsNullOrWhiteSpace(mainTaxFormula))
            {
                if (!BusinessValuesEqual(mainSheet.Cell(mainRow, 8), paymentSheet.Cell(paymentRow, 8)))
                {
                    throw new InvalidDataException("海南阶段二支付方工作表税率字段未与主汇总同步：" + key);
                }
            }
            else if (!string.Equals(
                TextUtil.S(paymentSheet.Cell(paymentRow, 8).FormulaA1),
                "J" + paymentRow + "-I" + paymentRow,
                StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidDataException("海南阶段二支付方工作表税率公式未与主汇总同步：" + key);
            }

            foreach (var column in new[] { 9, 10 })
            {
                var mainCell = mainSheet.Cell(mainRow, column);
                var paymentCell = paymentSheet.Cell(paymentRow, column);
                var equal = mainCell.DataType == XLDataType.Number
                    && paymentCell.DataType == XLDataType.Number
                    ? HainanStage2ExcelUtil.TaxRatesEqual(mainCell.GetDouble(), paymentCell.GetDouble())
                    : BusinessValuesEqual(mainCell, paymentCell);
                if (!equal)
                {
                    throw new InvalidDataException("海南阶段二支付方工作表税率字段未与主汇总同步：" + key);
                }
            }

            if (!BusinessValuesEqual(mainSheet.Cell(mainRow, 11), paymentSheet.Cell(paymentRow, 11)))
            {
                throw new InvalidDataException("海南阶段二支付方工作表负责人字段未与主汇总同步：" + key);
            }

            var mainCumulativeColumn = SummaryColumn(mainSheet, "累计代理费总计");
            var paymentCumulativeColumn = SummaryColumn(paymentSheet, "累计代理费总计");
            foreach (var offset in new[] { 1, 4, 5, 6, 7, 9 })
            {
                if (!BusinessValuesEqual(
                    mainSheet.Cell(mainRow, mainCumulativeColumn + offset),
                    paymentSheet.Cell(paymentRow, paymentCumulativeColumn + offset)))
                {
                    throw new InvalidDataException("海南阶段二支付方工作表借款、扣除或备注字段未与主汇总同步：" + key);
                }
            }
        }

        private static bool BusinessValuesEqual(IXLCell left, IXLCell right)
        {
            if (left.IsEmpty() || right.IsEmpty())
            {
                return left.IsEmpty() && right.IsEmpty();
            }

            if (left.DataType == XLDataType.Number && right.DataType == XLDataType.Number)
            {
                return AmountsEqual(left.GetDouble(), right.GetDouble());
            }

            if (left.DataType == XLDataType.DateTime && right.DataType == XLDataType.DateTime)
            {
                return left.GetDateTime() == right.GetDateTime();
            }

            return Stage2OpaqueText.AreEquivalent(left.GetFormattedString(), right.GetFormattedString());
        }

        private static bool AmountsEqual(double left, double right)
        {
            return Math.Abs(left - right) <= Stage2SettlementCalculator.AmountTolerance;
        }

        private static string SummaryKey(HainanStage2SummaryMetaRow row)
        {
            return HainanStage2ExcelUtil.SummaryKey(row.Entity, row.Kind);
        }

        private static double ParseMonthlyDeduction(string text)
        {
            var matchWan = Regex.Match(TextUtil.S(text), "每月扣除([0-9.]+)万");
            if (matchWan.Success)
            {
                return Convert.ToDouble(matchWan.Groups[1].Value);
            }

            var matchYuan = Regex.Match(TextUtil.S(text), "每月扣除([0-9.]+)元");
            if (matchYuan.Success)
            {
                return Math.Round(Convert.ToDouble(matchYuan.Groups[1].Value) / 10000, 4);
            }

            return 0;
        }

        private static string SumEverySix(int row, int firstColumn, int beforeColumn)
        {
            var parts = new List<string>();
            for (var column = firstColumn; column < beforeColumn; column += 6)
            {
                parts.Add(ClosedXmlUtil.ColumnLetter(column) + row);
            }

            return parts.Count == 0 ? "0" : string.Join("+", parts);
        }

        private static string SumFormula(int row, int startColumn, int endColumn)
        {
            var parts = new List<string>();
            for (var column = startColumn; column <= endColumn; column++)
            {
                parts.Add(ClosedXmlUtil.ColumnLetter(column) + row);
            }

            return string.Join("+", parts);
        }

        private static void SetNullableNumber(IXLCell cell, double? value)
        {
            if (value.HasValue)
            {
                cell.Value = value.Value;
            }
            else
            {
                cell.Clear(XLClearOptions.Contents);
            }
        }
    }
}
