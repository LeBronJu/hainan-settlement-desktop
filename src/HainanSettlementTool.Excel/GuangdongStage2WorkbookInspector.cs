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
    internal sealed class GuangdongStage2WorkbookInspector
    {
        private static readonly Regex PeriodRegex = new Regex(
            @"(?<year>\d{4})(?<yearSuffix>\s*年\s*)(?<month>\d{1,2})(?<monthSuffix>\s*月)",
            RegexOptions.Compiled);
        private static readonly Regex SettlementDateRegex = new Regex(
            @"(?<year>\d{4})(?<yearSuffix>\s*年\s*)(?<month>\d{1,2})(?<monthSuffix>\s*月\s*)(?<day>\d{1,2})(?<daySuffix>\s*日)",
            RegexOptions.Compiled);

        public IList<GuangdongStage2WorkbookPreparation> Analyze(GuangdongStage2MonthPreparationOptions options)
        {
            var preparations = new List<GuangdongStage2WorkbookPreparation>();
            AnalyzeDirectory(options, GuangdongStage2SettlementKinds.Proxy, options.ProxyDirectory, preparations);
            AnalyzeDirectory(options, GuangdongStage2SettlementKinds.Intermediary, options.IntermediaryDirectory, preparations);
            AnalyzeDirectory(options, GuangdongStage2SettlementKinds.Refund, options.RefundDirectory, preparations);
            return preparations;
        }

        private static void AnalyzeDirectory(
            GuangdongStage2MonthPreparationOptions options,
            string settlementKind,
            string root,
            ICollection<GuangdongStage2WorkbookPreparation> preparations)
        {
            if (string.IsNullOrWhiteSpace(root))
            {
                return;
            }

            foreach (var path in Directory.EnumerateFiles(root, "*.xlsx", SearchOption.AllDirectories)
                .Where(IsCandidateWorkbook)
                .OrderBy(item => item, StringComparer.OrdinalIgnoreCase))
            {
                preparations.Add(AnalyzeWorkbook(options, settlementKind, root, path));
            }
        }

        private static bool IsCandidateWorkbook(string path)
        {
            var name = Path.GetFileName(path);
            return !name.StartsWith("~$", StringComparison.Ordinal)
                && !name.StartsWith("._", StringComparison.Ordinal);
        }

        private static GuangdongStage2WorkbookPreparation AnalyzeWorkbook(
            GuangdongStage2MonthPreparationOptions options,
            string settlementKind,
            string root,
            string path)
        {
            var plan = NewPlan(options, settlementKind, root, path);
            try
            {
                using (var workbook = GuangdongStage2ExcelUtil.OpenWorkbookShared(path))
                {
                    var targetName = plan.TargetSheetName;
                    var sourceName = plan.SourceSheetName;
                    IXLWorksheet targetWorksheet;
                    IXLWorksheet sourceWorksheet;
                    var targetExists = workbook.TryGetWorksheet(targetName, out targetWorksheet);
                    var sourceExists = workbook.TryGetWorksheet(sourceName, out sourceWorksheet);
                    plan.TargetSheetExisted = targetExists;

                    if (!targetExists && !sourceExists)
                    {
                        return Skipped(plan, GuangdongStage2IssueKinds.MissingPreviousMonthSheet,
                            "缺少标准上月 sheet " + sourceName + "。非标准命名 sheet 不参与来源选择。");
                    }

                    var worksheet = targetExists ? targetWorksheet : sourceWorksheet;
                    var layout = InspectWorksheet(options, settlementKind, worksheet, targetExists, out var issueKind, out var message);
                    if (layout == null)
                    {
                        return Skipped(plan, issueKind, message);
                    }

                    plan.DetailRowCount = layout.DetailEndRow - layout.DetailStartRow + 1;
                    plan.PowerNeedsClearing = HasPowerContents(worksheet, layout, out issueKind, out message);
                    if (!string.IsNullOrWhiteSpace(issueKind))
                    {
                        return Skipped(plan, issueKind, message);
                    }

                    plan.TotalPowerNeedsReset = HasHardcodedTotalPower(worksheet, layout, out issueKind, out message);
                    if (!string.IsNullOrWhiteSpace(issueKind))
                    {
                        return Skipped(plan, issueKind, message);
                    }

                    plan.PeriodNeedsUpdate = worksheet.Cell(layout.PeriodCellAddress).GetFormattedString() != layout.TargetPeriodText;
                    plan.SettlementDateNeedsUpdate = worksheet.Cell(layout.SettlementDateCellAddress).GetFormattedString() != layout.TargetSettlementDateText;

                    if (!targetExists)
                    {
                        plan.Action = GuangdongStage2PreparationActions.CreateTargetMonth;
                        plan.Message = "将从标准 sheet " + sourceName + " 创建 " + targetName + "。";
                    }
                    else if (plan.PowerNeedsClearing || plan.PeriodNeedsUpdate || plan.SettlementDateNeedsUpdate || plan.TotalPowerNeedsReset)
                    {
                        plan.Action = GuangdongStage2PreparationActions.NormalizeExistingTargetMonth;
                        plan.Message = "将保留现有标准 sheet " + targetName + " 并整理电量和日期。";
                    }
                    else
                    {
                        plan.Action = GuangdongStage2PreparationActions.AlreadyPrepared;
                        plan.Message = "现有标准 sheet " + targetName + " 已准备完成。";
                    }

                    return new GuangdongStage2WorkbookPreparation { Plan = plan, Layout = layout };
                }
            }
            catch (Exception ex)
            {
                return Skipped(plan, GuangdongStage2IssueKinds.UnreadableWorkbook, "无法读取 workbook：" + ex.Message);
            }
        }

        private static GuangdongStage2WorkbookPlan NewPlan(
            GuangdongStage2MonthPreparationOptions options,
            string settlementKind,
            string root,
            string path)
        {
            return new GuangdongStage2WorkbookPlan
            {
                SettlementKind = settlementKind,
                SourceRoot = root,
                SourcePath = path,
                RelativePath = GuangdongStage2ExcelUtil.RelativePath(root, path),
                SourceSheetName = GuangdongStage2ExcelUtil.MonthSheetName(options.Month - 1),
                TargetSheetName = GuangdongStage2ExcelUtil.MonthSheetName(options.Month)
            };
        }

        private static GuangdongStage2WorkbookPreparation Skipped(
            GuangdongStage2WorkbookPlan plan,
            string issueKind,
            string message)
        {
            plan.Action = GuangdongStage2PreparationActions.Skipped;
            plan.IssueKind = issueKind;
            plan.Message = message;
            return new GuangdongStage2WorkbookPreparation { Plan = plan };
        }

        private static GuangdongStage2WorksheetLayout InspectWorksheet(
            GuangdongStage2MonthPreparationOptions options,
            string settlementKind,
            IXLWorksheet worksheet,
            bool targetExists,
            out string issueKind,
            out string message)
        {
            issueKind = null;
            message = null;
            var title = TextUtil.S(worksheet.Cell(1, 1).GetFormattedString());
            var expectedTitle = GuangdongStage2ExcelUtil.ExpectedTitle(settlementKind);
            if (!string.Equals(title, expectedTitle, StringComparison.Ordinal))
            {
                issueKind = GuangdongStage2IssueKinds.UnexpectedTitle;
                message = "标题不是预期的“" + expectedTitle + "”：" + title;
                return null;
            }

            var headerRow = FindPowerHeaderRow(worksheet);
            if (headerRow == 0)
            {
                issueKind = GuangdongStage2IssueKinds.InvalidPowerHeader;
                message = "未找到 C-F 的总实际电量、峰、平、谷标准表头。";
                return null;
            }

            var totalRow = FindTotalRow(worksheet, headerRow + 1);
            if (totalRow <= headerRow + 1)
            {
                issueKind = GuangdongStage2IssueKinds.MissingTotalRow;
                message = "未找到明细区之后的合计行。";
                return null;
            }

            var periodField = FindDateField(worksheet, "所属期", PeriodRegex, false);
            if (periodField == null)
            {
                issueKind = GuangdongStage2IssueKinds.MissingPeriod;
                message = "未找到可识别的所属期。";
                return null;
            }

            var settlementDateField = FindDateField(worksheet, "结算日期", SettlementDateRegex, true);
            if (settlementDateField == null)
            {
                issueKind = GuangdongStage2IssueKinds.MissingSettlementDate;
                message = "未找到可识别的结算日期。";
                return null;
            }

            var expectedPeriod = new DateTime(options.Year, options.Month, 1);
            var previousPeriod = expectedPeriod.AddMonths(-1);
            if ((!targetExists && !SameMonth(periodField.Value, previousPeriod))
                || (targetExists && !SameMonth(periodField.Value, expectedPeriod) && !SameMonth(periodField.Value, previousPeriod)))
            {
                issueKind = GuangdongStage2IssueKinds.InvalidPeriod;
                message = "所属期与标准 sheet 名不一致：" + periodField.Text;
                return null;
            }

            var desiredSettlementMonth = expectedPeriod.AddMonths(1);
            DateTime targetSettlementDate;
            if (!targetExists)
            {
                targetSettlementDate = settlementDateField.Value.AddMonths(1);
                if (!SameMonth(targetSettlementDate, desiredSettlementMonth))
                {
                    issueKind = GuangdongStage2IssueKinds.InvalidSettlementDate;
                    message = "上月结算日期顺延后不是目标结算月份：" + settlementDateField.Text;
                    return null;
                }
            }
            else if (SameMonth(settlementDateField.Value, desiredSettlementMonth))
            {
                targetSettlementDate = settlementDateField.Value;
            }
            else if (SameMonth(settlementDateField.Value, expectedPeriod))
            {
                targetSettlementDate = settlementDateField.Value.AddMonths(1);
            }
            else
            {
                issueKind = GuangdongStage2IssueKinds.InvalidSettlementDate;
                message = "现有目标月 sheet 的结算日期既不是目标状态，也不是可顺延的上月状态：" + settlementDateField.Text;
                return null;
            }

            return new GuangdongStage2WorksheetLayout
            {
                HeaderRow = headerRow,
                DetailStartRow = headerRow + 1,
                DetailEndRow = totalRow - 1,
                TotalRow = totalRow,
                PeriodCellAddress = periodField.CellAddress,
                SettlementDateCellAddress = settlementDateField.CellAddress,
                TargetPeriodText = ReplacePeriod(periodField.Text, expectedPeriod),
                TargetSettlementDateText = ReplaceSettlementDate(settlementDateField.Text, targetSettlementDate)
            };
        }

        private static int FindPowerHeaderRow(IXLWorksheet worksheet)
        {
            var last = Math.Min(10, worksheet.LastRowUsed()?.RowNumber() ?? 10);
            for (var row = 1; row <= last; row++)
            {
                if (TextUtil.S(worksheet.Cell(row, 3).GetFormattedString()).Contains("总实际电量")
                    && TextUtil.S(worksheet.Cell(row, 4).GetFormattedString()) == "峰"
                    && TextUtil.S(worksheet.Cell(row, 5).GetFormattedString()) == "平"
                    && TextUtil.S(worksheet.Cell(row, 6).GetFormattedString()) == "谷")
                {
                    return row;
                }
            }

            return 0;
        }

        private static int FindTotalRow(IXLWorksheet worksheet, int startRow)
        {
            var last = worksheet.LastRowUsed()?.RowNumber() ?? startRow;
            for (var row = startRow; row <= last; row++)
            {
                if (TextUtil.S(worksheet.Cell(row, 1).GetFormattedString()) == "合计")
                {
                    return row;
                }
            }

            return 0;
        }

        private static GuangdongStage2DateField FindDateField(
            IXLWorksheet worksheet,
            string label,
            Regex regex,
            bool includeDay)
        {
            var matches = new List<GuangdongStage2DateField>();
            var lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? 1;
            var lastRow = Math.Min(10, worksheet.LastRowUsed()?.RowNumber() ?? 10);
            for (var row = 1; row <= lastRow; row++)
            {
                for (var column = 1; column <= lastColumn; column++)
                {
                    var cell = worksheet.Cell(row, column);
                    var text = cell.GetFormattedString();
                    if (text.IndexOf(label, StringComparison.Ordinal) < 0)
                    {
                        continue;
                    }

                    var match = regex.Match(text);
                    if (!match.Success)
                    {
                        continue;
                    }

                    var year = int.Parse(match.Groups["year"].Value, CultureInfo.InvariantCulture);
                    var month = int.Parse(match.Groups["month"].Value, CultureInfo.InvariantCulture);
                    var day = includeDay
                        ? int.Parse(match.Groups["day"].Value, CultureInfo.InvariantCulture)
                        : 1;
                    DateTime value;
                    try
                    {
                        value = new DateTime(year, month, day);
                    }
                    catch (ArgumentOutOfRangeException)
                    {
                        continue;
                    }

                    matches.Add(new GuangdongStage2DateField
                    {
                        CellAddress = cell.Address.ToStringRelative(),
                        Text = text,
                        Value = value
                    });
                }
            }

            return matches.Count == 1 ? matches[0] : null;
        }

        private static bool HasPowerContents(
            IXLWorksheet worksheet,
            GuangdongStage2WorksheetLayout layout,
            out string issueKind,
            out string message)
        {
            issueKind = null;
            message = null;
            var hasContents = false;
            for (var row = layout.DetailStartRow; row <= layout.DetailEndRow; row++)
            {
                for (var column = 3; column <= 6; column++)
                {
                    var cell = worksheet.Cell(row, column);
                    if (cell.HasFormula)
                    {
                        issueKind = GuangdongStage2IssueKinds.UnexpectedPowerFormula;
                        message = "明细电量区域存在公式，未自动清空：" + cell.Address.ToStringRelative();
                        return false;
                    }

                    hasContents = hasContents || !cell.IsEmpty();
                }
            }

            return hasContents;
        }

        private static bool HasHardcodedTotalPower(
            IXLWorksheet worksheet,
            GuangdongStage2WorksheetLayout layout,
            out string issueKind,
            out string message)
        {
            issueKind = null;
            message = null;
            var needsReset = false;
            for (var column = 3; column <= 6; column++)
            {
                var cell = worksheet.Cell(layout.TotalRow, column);
                if (cell.HasFormula || cell.IsEmpty())
                {
                    continue;
                }

                double value;
                if (!double.TryParse(cell.GetFormattedString(), NumberStyles.Any, CultureInfo.InvariantCulture, out value)
                    && !double.TryParse(cell.GetFormattedString(), NumberStyles.Any, CultureInfo.CurrentCulture, out value))
                {
                    issueKind = GuangdongStage2IssueKinds.InvalidPowerHeader;
                    message = "合计行 C-F 包含无法识别的非公式值：" + cell.Address.ToStringRelative();
                    return false;
                }

                needsReset = needsReset || Math.Abs(value) > 0.0000001;
            }

            return needsReset;
        }

        private static bool SameMonth(DateTime left, DateTime right)
        {
            return left.Year == right.Year && left.Month == right.Month;
        }

        private static string ReplacePeriod(string original, DateTime value)
        {
            var match = PeriodRegex.Match(original);
            return ReplaceDateMatch(original, match, value, false);
        }

        private static string ReplaceSettlementDate(string original, DateTime value)
        {
            var match = SettlementDateRegex.Match(original);
            return ReplaceDateMatch(original, match, value, true);
        }

        private static string ReplaceDateMatch(string original, Match match, DateTime value, bool includeDay)
        {
            var monthFormat = match.Groups["month"].Value.Length >= 2 ? "00" : "0";
            var replacement = value.Year.ToString(CultureInfo.InvariantCulture)
                + match.Groups["yearSuffix"].Value
                + value.Month.ToString(monthFormat, CultureInfo.InvariantCulture)
                + match.Groups["monthSuffix"].Value;
            if (includeDay)
            {
                var dayFormat = match.Groups["day"].Value.Length >= 2 ? "00" : "0";
                replacement += value.Day.ToString(dayFormat, CultureInfo.InvariantCulture)
                    + match.Groups["daySuffix"].Value;
            }

            return original.Substring(0, match.Index)
                + replacement
                + original.Substring(match.Index + match.Length);
        }
    }
}
