using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using HainanSettlementTool.Core.Models;
using HainanSettlementTool.Core.Services;

namespace HainanSettlementTool.Excel
{
    internal static class ChongqingStage2SummaryWorkbookWriter
    {
        public static string BuildSummary(ChongqingStage2Options options, IList<GroupSettlementTotal> groups, IList<string> warnings)
        {
            var outputName = string.IsNullOrWhiteSpace(options.OutputSummaryName)
                ? "【2026年重庆代理费汇总表-" + options.Month + "月自动化】.xlsx"
                : options.OutputSummaryName;
            var outputPath = Path.Combine(options.OutputDirectory, outputName);
            FileAccessGuard.RequireWritableWorkbook(outputPath, "重庆阶段二输出汇总表");
            ChongqingStage2ExcelUtil.CopyWorkbookShared(options.SummaryTemplatePath, outputPath, overwrite: true);

            using (var workbook = new XLWorkbook(outputPath))
            {
                var mainSheet = FindSummaryWorksheet(workbook);
                var mainMeta = ReadSummaryMeta(mainSheet);
                var paymentPartyByKey = BuildPaymentPartyIndex(options, groups, mainMeta);
                WriteSummarySheet(mainSheet, groups, options.Month, null, warnings, paymentPartyByKey);

                foreach (var paymentParty in ChongqingStage2PaymentParties.Supported)
                {
                    var sheet = PreparePaymentPartySheet(workbook, paymentParty, options.Month, mainSheet);
                    var allowedKeys = new HashSet<string>(paymentPartyByKey
                        .Where(item => item.Value == paymentParty)
                        .Select(item => item.Key));
                    var partyGroups = groups
                        .Where(group => PartyForSummaryTotal(group, paymentPartyByKey) == paymentParty)
                        .ToList();
                    foreach (var group in partyGroups)
                    {
                        allowedKeys.Add(ChongqingStage2Keys.SummaryKey(group.Entity, group.Kind));
                    }

                    WriteSummarySheet(sheet, partyGroups, options.Month, allowedKeys, warnings, paymentPartyByKey);
                }

                ChongqingStage2ExcelUtil.SaveWorkbook(workbook, outputPath);
            }

            return outputPath;
        }

        public static void AddSummaryPaymentIssues(
            ChongqingStage2Options options,
            IList<GroupSettlementTotal> groups,
            IList<ChongqingStage2CheckIssue> issues)
        {
            if (groups.Count == 0)
            {
                return;
            }

            using (var stream = File.Open(options.SummaryTemplatePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var workbook = new XLWorkbook(stream))
            {
                var mainSheet = FindSummaryWorksheet(workbook);
                var knownKeys = new HashSet<string>(ReadSummaryMeta(mainSheet).Select(item => ChongqingStage2Keys.SummaryKey(item.Entity, item.Kind)));
                foreach (var group in groups)
                {
                    if (knownKeys.Contains(ChongqingStage2Keys.SummaryKey(group.Entity, group.Kind))
                        || HasPaymentDecision(options, group.Entity, group.Kind))
                    {
                        continue;
                    }

                    var issue = new ChongqingStage2CheckIssue
                    {
                        Severity = "确认",
                        Category = "新增汇总主体支付方选择",
                        Kind = ChongqingStage2IssueKinds.NewSummarySubjectPaymentPartyRequired,
                        SettlementKind = group.Kind,
                        Owner = group.Owner,
                        Entity = group.Entity,
                        TemplateFile = options.SummaryTemplatePath,
                        SheetName = mainSheet.Name,
                        Message = "重庆汇总表模板缺少" + group.Kind + "主体“" + group.Entity + "”，需要选择支付方后才能生成支付方月度 sheet。",
                        Suggestion = "请选择清能或清辉；本次选择只用于当前输出汇总表副本。",
                        RequiresPaymentPartySelection = true
                    };
                    issue.AvailablePaymentParties.AddRange(ChongqingStage2PaymentParties.Supported);
                    issues.Add(issue);
                }
            }
        }

        private static void WriteSummarySheet(
            IXLWorksheet worksheet,
            IList<GroupSettlementTotal> groups,
            int month,
            ISet<string> allowedKeys,
            IList<string> warnings,
            IDictionary<string, string> paymentPartyByKey)
        {
            DeleteSummaryRowsNotAllowed(worksheet, allowedKeys);
            var monthColumn = FindOrInsertSummaryMonthBlock(worksheet, month);
            var cumulativeColumn = SummaryColumn(worksheet, "当年费用总计");
            if (allowedKeys != null)
            {
                ApplyPaymentPartySheetVisibility(worksheet, monthColumn, cumulativeColumn);
            }

            var totalByKey = groups.ToDictionary(group => ChongqingStage2Keys.SummaryKey(group.Entity, group.Kind), group => group);
            var knownKeys = new HashSet<string>(ReadSummaryMeta(worksheet).Select(row => ChongqingStage2Keys.SummaryKey(row.Entity, row.Kind)));
            var newGroups = groups.Where(group => !knownKeys.Contains(ChongqingStage2Keys.SummaryKey(group.Entity, group.Kind))).ToList();
            var newKeys = new HashSet<string>(newGroups.Select(group => ChongqingStage2Keys.SummaryKey(group.Entity, group.Kind)));
            InsertNewSummaryRows(worksheet, newGroups, warnings, paymentPartyByKey);

            var rows = ReadSummaryMeta(worksheet).OrderBy(row => row.Row).ToList();
            for (var index = 0; index < rows.Count; index++)
            {
                var row = rows[index].Row;
                var info = rows[index];
                GroupSettlementTotal total;
                totalByKey.TryGetValue(ChongqingStage2Keys.SummaryKey(info.Entity, info.Kind), out total);
                worksheet.Cell(row, 1).Value = index + 1;
                WriteSummaryValues(worksheet, row, monthColumn, cumulativeColumn, total, info.Entity, info.Kind, month, newKeys.Contains(ChongqingStage2Keys.SummaryKey(info.Entity, info.Kind)), paymentPartyByKey);
            }

            var totalRow = FindSummaryTotalRow(worksheet);
            worksheet.Cell(totalRow, 1).Value = "合计";
            WriteSummaryTotalRow(worksheet, totalRow, cumulativeColumn);
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
            IDictionary<string, string> paymentPartyByKey)
        {
            worksheet.Cell(row, monthColumn).Clear(XLClearOptions.Contents);
            worksheet.Cell(row, monthColumn + 1).Clear(XLClearOptions.Contents);
            worksheet.Cell(row, monthColumn + 2).Clear(XLClearOptions.Contents);

            if (total != null)
            {
                worksheet.Cell(row, monthColumn + SummaryKindOffset(total.Kind)).Value = Math.Round(total.ExpectedNet, 4);
            }

            worksheet.Cell(row, monthColumn + 3).FormulaA1 = SumFormula(row, monthColumn, monthColumn + 2);
            worksheet.Cell(row, monthColumn + 5).FormulaA1 = ClosedXmlUtil.ColumnLetter(monthColumn + 3) + row + "-" + ClosedXmlUtil.ColumnLetter(monthColumn + 4) + row;

            if (isNewRow)
            {
                var firstMonthColumn = FirstSummaryMonthColumn(worksheet);
                worksheet.Cell(row, cumulativeColumn).FormulaA1 = SumEverySix(row, firstMonthColumn, cumulativeColumn - 1, 3);
                worksheet.Cell(row, cumulativeColumn + 2).FormulaA1 = SumEverySix(row, firstMonthColumn, cumulativeColumn - 1, 4);
                worksheet.Cell(row, cumulativeColumn + 3).FormulaA1 = ClosedXmlUtil.ColumnLetter(cumulativeColumn + 1) + row + "-" + ClosedXmlUtil.ColumnLetter(cumulativeColumn + 2) + row;
                worksheet.Cell(row, cumulativeColumn + 6).Value = new DateTime(ChongqingStage2Layout.Year, month, 1);
                worksheet.Cell(row, cumulativeColumn + 7).Value = PaymentPartyFromIndex(paymentPartyByKey, entity, kind);
            }
        }

        private static void InsertNewSummaryRows(
            IXLWorksheet worksheet,
            IList<GroupSettlementTotal> newGroups,
            IList<string> warnings,
            IDictionary<string, string> paymentPartyByKey)
        {
            if (newGroups.Count == 0)
            {
                return;
            }

            var totalRow = FindSummaryTotalRow(worksheet);
            var templateRow = Math.Max(ChongqingStage2Layout.SummaryDataStartRow, totalRow - 1);
            foreach (var group in newGroups)
            {
                worksheet.Row(totalRow).InsertRowsAbove(1);
                worksheet.Row(templateRow).CopyTo(worksheet.Row(totalRow));
                worksheet.Row(totalRow).Clear(XLClearOptions.Contents);
                worksheet.Cell(totalRow, 1).Value = totalRow - ChongqingStage2Layout.SummaryDataStartRow + 1;
                worksheet.Cell(totalRow, 2).Value = group.Entity;
                worksheet.Cell(totalRow, 3).Value = group.Kind;
                worksheet.Cell(totalRow, 6).Value = group.Entity;
                worksheet.Cell(totalRow, 11).Value = group.Owner;
                var cumulativeColumn = SummaryColumn(worksheet, "当年费用总计");
                worksheet.Cell(totalRow, cumulativeColumn + 7).Value = PaymentPartyFromIndex(paymentPartyByKey, group.Entity, group.Kind);
                warnings.Add("新增重庆汇总主体：" + group.Kind + " " + group.Entity + "（负责人：" + group.Owner + "；支付方：" + PaymentPartyFromIndex(paymentPartyByKey, group.Entity, group.Kind) + "），仅自动填入最小必要字段，请人工复核收款、发票和借支字段。");
                totalRow++;
            }
        }

        private static IDictionary<string, string> BuildPaymentPartyIndex(
            ChongqingStage2Options options,
            IList<GroupSettlementTotal> groups,
            IList<ChongqingSummaryMetaRow> mainMeta)
        {
            var partyByKey = new Dictionary<string, string>();
            foreach (var row in mainMeta)
            {
                var key = ChongqingStage2Keys.SummaryKey(row.Entity, row.Kind);
                if (!string.IsNullOrWhiteSpace(row.PaymentParty))
                {
                    partyByKey[key] = row.PaymentParty;
                }
            }

            foreach (var group in groups)
            {
                var key = ChongqingStage2Keys.SummaryKey(group.Entity, group.Kind);
                if (partyByKey.ContainsKey(key))
                {
                    continue;
                }

                string paymentParty;
                if (!TryGetPaymentPartyDecision(options, group.Entity, group.Kind, out paymentParty))
                {
                    throw new InvalidOperationException("重庆阶段二新增汇总主体支付方未选择：" + group.Kind + " " + group.Entity + "。请在预检窗口选择清能或清辉后再生成。");
                }

                partyByKey[key] = paymentParty;
            }

            return partyByKey;
        }

        private static IXLWorksheet PreparePaymentPartySheet(XLWorkbook workbook, string paymentParty, int month, IXLWorksheet mainSheet)
        {
            var targetName = paymentParty + month + "月";
            var existing = workbook.Worksheets.FirstOrDefault(sheet => sheet.Name == targetName);
            var source = PreviousPaymentPartySheet(workbook, paymentParty, month);
            if (existing != null)
            {
                if (source == null)
                {
                    return existing;
                }

                existing.Delete();
            }

            if (source != null)
            {
                return source.CopyTo(targetName);
            }

            return mainSheet.CopyTo(targetName);
        }

        private static IXLWorksheet PreviousPaymentPartySheet(XLWorkbook workbook, string paymentParty, int month)
        {
            return workbook.Worksheets
                .Select(sheet =>
                {
                    int parsedMonth;
                    return new { Sheet = sheet, Matched = TryParsePaymentPartySheet(sheet.Name, paymentParty, out parsedMonth), Month = parsedMonth };
                })
                .Where(item => item.Matched && item.Month < month)
                .OrderBy(item => item.Month)
                .Select(item => item.Sheet)
                .LastOrDefault();
        }

        private static bool TryParsePaymentPartySheet(string name, string paymentParty, out int month)
        {
            month = 0;
            var match = Regex.Match(TextUtil.S(name), "^" + Regex.Escape(paymentParty) + "(\\d{1,2})月$");
            return match.Success && int.TryParse(match.Groups[1].Value, out month);
        }

        private static int FindOrInsertSummaryMonthBlock(IXLWorksheet worksheet, int month)
        {
            var existing = FindSummaryMonthBlock(worksheet, month);
            if (existing > 0)
            {
                return existing;
            }

            var cumulativeColumn = SummaryColumn(worksheet, "当年费用总计");
            var insertAt = cumulativeColumn;
            worksheet.Column(insertAt).InsertColumnsBefore(6);
            for (var offset = 0; offset < 6; offset++)
            {
                var sourceColumn = Math.Max(1, insertAt - 6 + offset);
                worksheet.Column(sourceColumn).CopyTo(worksheet.Column(insertAt + offset));
                worksheet.Column(insertAt + offset).Unhide();
            }

            worksheet.Cell(2, insertAt).Value = new DateTime(ChongqingStage2Layout.Year, month, 1);
            worksheet.Cell(2, insertAt + 5).Value = "当月实际支付";
            worksheet.Cell(3, insertAt).Value = "代理费";
            worksheet.Cell(3, insertAt + 1).Value = "居间费";
            worksheet.Cell(3, insertAt + 2).Value = "退补电费";
            worksheet.Cell(3, insertAt + 3).Value = "费用合计";
            worksheet.Cell(3, insertAt + 4).Value = "当月抵扣";
            worksheet.Cell(3, insertAt + 5).Clear(XLClearOptions.Contents);
            return insertAt;
        }

        private static void ApplyPaymentPartySheetVisibility(IXLWorksheet worksheet, int monthColumn, int cumulativeColumn)
        {
            var lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? cumulativeColumn + 8;
            for (var column = 1; column <= lastColumn; column++)
            {
                if (ChongqingStage2ExcelUtil.CellText(worksheet.Cell(3, column)) != "代理费")
                {
                    continue;
                }

                for (var offset = 0; offset < 6 && column + offset <= lastColumn; offset++)
                {
                    worksheet.Column(column + offset).Hide();
                }
            }

            for (var offset = 0; offset < 6 && monthColumn + offset <= lastColumn; offset++)
            {
                worksheet.Column(monthColumn + offset).Unhide();
            }

            for (var column = cumulativeColumn; column <= Math.Min(lastColumn, cumulativeColumn + 6); column++)
            {
                worksheet.Column(column).Unhide();
            }

            if (cumulativeColumn + 7 <= lastColumn)
            {
                worksheet.Column(cumulativeColumn + 7).Hide();
            }

            if (cumulativeColumn + 8 <= lastColumn)
            {
                worksheet.Column(cumulativeColumn + 8).Unhide();
            }
        }

        private static int FindSummaryMonthBlock(IXLWorksheet worksheet, int month)
        {
            var lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? 1;
            for (var column = 1; column <= lastColumn; column++)
            {
                if (ChongqingStage2ExcelUtil.CellText(worksheet.Cell(3, column)) != "代理费")
                {
                    continue;
                }

                if (SummaryHeaderMatchesMonth(worksheet.Cell(2, column), month))
                {
                    return column;
                }
            }

            return 0;
        }

        private static bool SummaryHeaderMatchesMonth(IXLCell cell, int month)
        {
            try
            {
                if (cell.DataType == XLDataType.DateTime && cell.GetDateTime().Month == month && cell.GetDateTime().Year == ChongqingStage2Layout.Year)
                {
                    return true;
                }
            }
            catch
            {
            }

            var text = ChongqingStage2ExcelUtil.CellText(cell);
            return text.Contains(ChongqingStage2Layout.Year.ToString(CultureInfo.InvariantCulture))
                && (text.Contains(month + "月") || text.Contains("-" + month.ToString("00", CultureInfo.InvariantCulture) + "-"));
        }

        private static void DeleteSummaryRowsNotAllowed(IXLWorksheet worksheet, ISet<string> allowedKeys)
        {
            if (allowedKeys == null)
            {
                return;
            }

            var totalRow = FindSummaryTotalRow(worksheet);
            for (var row = totalRow - 1; row >= ChongqingStage2Layout.SummaryDataStartRow; row--)
            {
                var entity = ChongqingStage2ExcelUtil.CellText(worksheet.Cell(row, 2));
                var kind = ChongqingStage2ExcelUtil.CellText(worksheet.Cell(row, 3));
                if (string.IsNullOrWhiteSpace(entity) || !allowedKeys.Contains(ChongqingStage2Keys.SummaryKey(entity, kind)))
                {
                    worksheet.Row(row).Delete();
                }
            }
        }

        private static void WriteSummaryTotalRow(IXLWorksheet worksheet, int totalRow, int cumulativeColumn)
        {
            var firstMonthColumn = FirstSummaryMonthColumn(worksheet);
            for (var column = firstMonthColumn; column <= cumulativeColumn + 3; column++)
            {
                var letter = ClosedXmlUtil.ColumnLetter(column);
                worksheet.Cell(totalRow, column).FormulaA1 = "SUM(" + letter + ChongqingStage2Layout.SummaryDataStartRow + ":" + letter + (totalRow - 1) + ")";
            }
        }

        private static int FirstSummaryMonthColumn(IXLWorksheet worksheet)
        {
            var lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? 1;
            for (var column = 1; column <= lastColumn; column++)
            {
                if (ChongqingStage2ExcelUtil.CellText(worksheet.Cell(3, column)) == "代理费")
                {
                    return column;
                }
            }

            throw new InvalidOperationException(worksheet.Name + " 未找到重庆汇总表月度费用列。");
        }

        private static string SumEverySix(int row, int firstColumn, int lastColumn, int offset)
        {
            var parts = new List<string>();
            for (var column = firstColumn + offset; column <= lastColumn; column += 6)
            {
                parts.Add(ClosedXmlUtil.ColumnLetter(column) + row);
            }

            return string.Join("+", parts);
        }

        private static string SumFormula(int row, int firstColumn, int lastColumn)
        {
            var parts = new List<string>();
            for (var column = firstColumn; column <= lastColumn; column++)
            {
                parts.Add(ClosedXmlUtil.ColumnLetter(column) + row);
            }

            return string.Join("+", parts);
        }

        private static int SummaryKindOffset(string kind)
        {
            if (kind == ChongqingStage2SettlementKinds.Proxy)
            {
                return 0;
            }

            if (kind == ChongqingStage2SettlementKinds.Intermediary)
            {
                return 1;
            }

            return 2;
        }

        private static bool HasPaymentDecision(ChongqingStage2Options options, string entity, string kind)
        {
            string paymentParty;
            return TryGetPaymentPartyDecision(options, entity, kind, out paymentParty);
        }

        private static bool TryGetPaymentPartyDecision(ChongqingStage2Options options, string entity, string kind, out string paymentParty)
        {
            paymentParty = null;
            var key = ChongqingStage2Keys.SummaryKey(entity, kind);
            var decision = options.SummarySubjectDecisions
                .Where(item => item != null)
                .FirstOrDefault(item => ChongqingStage2Keys.SummaryKey(item.Entity, item.SettlementKind) == key);
            if (decision == null || string.IsNullOrWhiteSpace(decision.PaymentParty))
            {
                return false;
            }

            paymentParty = decision.PaymentParty;
            return true;
        }

        private static string PartyForSummaryTotal(GroupSettlementTotal total, IDictionary<string, string> partyByKey)
        {
            string party;
            if (partyByKey.TryGetValue(ChongqingStage2Keys.SummaryKey(total.Entity, total.Kind), out party))
            {
                return party;
            }

            throw new InvalidOperationException("重庆阶段二新增汇总主体支付方未选择：" + total.Kind + " " + total.Entity + "。");
        }

        private static string PaymentPartyFromIndex(IDictionary<string, string> paymentPartyByKey, string entity, string kind)
        {
            string paymentParty;
            if (paymentPartyByKey.TryGetValue(ChongqingStage2Keys.SummaryKey(entity, kind), out paymentParty))
            {
                return paymentParty;
            }

            throw new InvalidOperationException("重庆阶段二新增汇总主体支付方未选择：" + kind + " " + entity + "。");
        }

        private static IXLWorksheet FindSummaryWorksheet(XLWorkbook workbook)
        {
            var exact = workbook.Worksheets.FirstOrDefault(sheet => ChongqingStage2ExcelUtil.CellText(sheet.Cell("A1")).Contains("汇总")
                || sheet.Name == "汇总表");
            if (exact != null)
            {
                return exact;
            }

            return workbook.Worksheets.First();
        }

        private static List<ChongqingSummaryMetaRow> ReadSummaryMeta(IXLWorksheet worksheet)
        {
            var rows = new List<ChongqingSummaryMetaRow>();
            var totalRow = FindSummaryTotalRow(worksheet);
            var paymentPartyColumn = FindHeaderColumnInRows(worksheet, "支付方", 1, 3);
            for (var row = ChongqingStage2Layout.SummaryDataStartRow; row < totalRow; row++)
            {
                var entity = ChongqingStage2ExcelUtil.CellText(worksheet.Cell(row, 2));
                var kind = ChongqingStage2ExcelUtil.CellText(worksheet.Cell(row, 3));
                if (string.IsNullOrWhiteSpace(entity) || string.IsNullOrWhiteSpace(kind))
                {
                    continue;
                }

                rows.Add(new ChongqingSummaryMetaRow
                {
                    Row = row,
                    Entity = entity,
                    Kind = kind,
                    PaymentParty = paymentPartyColumn > 0 ? ChongqingStage2ExcelUtil.CellText(worksheet.Cell(row, paymentPartyColumn)) : string.Empty
                });
            }

            return rows;
        }

        private static int FindSummaryTotalRow(IXLWorksheet worksheet)
        {
            var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? ChongqingStage2Layout.SummaryDataStartRow;
            for (var row = ChongqingStage2Layout.SummaryDataStartRow; row <= lastRow; row++)
            {
                if (ChongqingStage2ExcelUtil.CellText(worksheet.Cell(row, 1)) == "合计")
                {
                    return row;
                }
            }

            return lastRow + 1;
        }

        private static int SummaryColumn(IXLWorksheet worksheet, string header)
        {
            var column = FindHeaderColumnInRows(worksheet, header, 1, 3);
            if (column <= 0)
            {
                throw new InvalidOperationException(worksheet.Name + " 未找到重庆汇总表表头“" + header + "”。");
            }

            return column;
        }

        private static int FindHeaderColumnInRows(IXLWorksheet worksheet, string header, int firstRow, int lastRow)
        {
            var lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? 1;
            for (var row = firstRow; row <= lastRow; row++)
            {
                for (var column = 1; column <= lastColumn; column++)
                {
                    if (ChongqingStage2ExcelUtil.CellText(worksheet.Cell(row, column)) == header)
                    {
                        return column;
                    }
                }
            }

            return 0;
        }
    }
}
