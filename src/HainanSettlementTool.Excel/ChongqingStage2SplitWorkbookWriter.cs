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
    internal static class ChongqingStage2SplitWorkbookWriter
    {
        public static void VerifyGeneratedSplitWorkbooks(
            IList<GroupSettlementTotal> groups,
            IList<ChongqingSettlementDetail> details,
            int month)
        {
            foreach (var group in groups)
            {
                var key = ChongqingStage2Keys.SummaryKey(group.Entity, group.Kind);
                var expectedRows = details
                    .Where(detail => ChongqingStage2Keys.SummaryKey(detail.Entity, detail.Kind) == key)
                    .ToList();
                using (var stream = File.Open(group.OutputFile, FileMode.Open, FileAccess.Read, FileShare.Read))
                using (var workbook = new XLWorkbook(stream))
                {
                    var sheetName = ChongqingStage2ExcelUtil.MonthSheetName(month);
                    var worksheet = workbook.Worksheets.FirstOrDefault(sheet => sheet.Name == sheetName);
                    if (worksheet == null)
                    {
                        throw new InvalidDataException("重庆阶段二分表缺少目标月份工作表：" + group.OutputFile);
                    }

                    var entity = ExtractEntityFromSplitSheet(worksheet);
                    if (TextUtil.CustomerKey(entity) != TextUtil.CustomerKey(group.Entity))
                    {
                        throw new InvalidDataException("重庆阶段二分表主体与分组不一致：" + group.OutputFile);
                    }

                    var detailRows = 0;
                    var totalRow = FindSplitTotalRow(worksheet);
                    for (var row = ChongqingStage2Layout.DetailDataStartRow; row < totalRow; row++)
                    {
                        if (!string.IsNullOrWhiteSpace(ChongqingStage2ExcelUtil.CellText(worksheet.Cell(row, 2))))
                        {
                            detailRows++;
                        }
                    }

                    if (detailRows != group.Rows || detailRows != expectedRows.Count)
                    {
                        throw new InvalidDataException("重庆阶段二分表明细行数与分组不一致：" + group.OutputFile);
                    }

                    for (var index = 0; index < expectedRows.Count; index++)
                    {
                        VerifyGeneratedDetailRow(
                            worksheet,
                            ChongqingStage2Layout.DetailDataStartRow + index,
                            expectedRows[index],
                            group.OutputFile);
                    }

                    var netColumn = group.Kind == ChongqingStage2SettlementKinds.Refund ? 15 : 14;
                    VerifyGeneratedNumber(
                        worksheet.Cell(totalRow, netColumn),
                        Math.Round(group.ExpectedNet, 4),
                        Stage2SettlementCalculator.AmountTolerance,
                        "合计净额",
                        group.OutputFile);
                }
            }
        }

        private static void VerifyGeneratedDetailRow(
            IXLWorksheet worksheet,
            int row,
            ChongqingSettlementDetail expected,
            string outputPath)
        {
            var actualCustomer = ChongqingStage2ExcelUtil.CellText(worksheet.Cell(row, 2));
            if (TextUtil.CustomerKey(actualCustomer) != TextUtil.CustomerKey(expected.Customer))
            {
                throw new InvalidDataException(
                    "重庆阶段二分表客户与台账分组不一致："
                    + outputPath + "，第" + row + "行。");
            }

            VerifyGeneratedNumber(worksheet.Cell(row, 4), Math.Round(expected.Sharp, 4), 0.00005d, "尖峰/尖电量", outputPath);
            VerifyGeneratedNumber(worksheet.Cell(row, 5), Math.Round(expected.Peak, 4), 0.00005d, "峰电量", outputPath);
            VerifyGeneratedNumber(worksheet.Cell(row, 6), Math.Round(expected.Flat, 4), 0.00005d, "平电量", outputPath);
            VerifyGeneratedNumber(worksheet.Cell(row, 7), Math.Round(expected.Valley, 4), 0.00005d, "谷电量", outputPath);
            VerifyGeneratedNumber(worksheet.Cell(row, 8), Math.Round(expected.Ratio, 6), 0.0000005d, "电量比例", outputPath);

            if (expected.Kind == ChongqingStage2SettlementKinds.Refund)
            {
                VerifyGeneratedNumber(worksheet.Cell(row, 9), Math.Round(expected.RefundSharpPrice, 6), 0.0000005d, "尖峰单价", outputPath);
                VerifyGeneratedNumber(worksheet.Cell(row, 10), Math.Round(expected.RefundPeakPrice, 6), 0.0000005d, "峰单价", outputPath);
                VerifyGeneratedNumber(worksheet.Cell(row, 11), Math.Round(expected.RefundFlatPrice, 6), 0.0000005d, "平单价", outputPath);
                VerifyGeneratedNumber(worksheet.Cell(row, 12), Math.Round(expected.RefundValleyPrice, 6), 0.0000005d, "谷单价", outputPath);
                VerifyGeneratedNumber(
                    worksheet.Cell(row, 16),
                    Math.Round(expected.TaxRate, 10),
                    ChongqingStage2ExcelUtil.TaxRateTolerance,
                    "扣税率",
                    outputPath);
                VerifyGeneratedFormula(worksheet.Cell(row, 15), "M" + row + "-N" + row, "实际收益", outputPath);
                VerifyGeneratedNumber(worksheet.Cell(row, 15), Math.Round(expected.CalculatedNet, 4), Stage2SettlementCalculator.AmountTolerance, "实际收益", outputPath);
                return;
            }

            VerifyGeneratedNumber(worksheet.Cell(row, 3), Math.Round(expected.Total, 4), 0.00005d, "总实际电量", outputPath);
            VerifyGeneratedNumber(worksheet.Cell(row, 9), Math.Round(expected.UnitPrice, 6), 0.0000005d, "利润单价", outputPath);
            VerifyGeneratedNumber(worksheet.Cell(row, 12), Math.Round(expected.RecoverShortfall, 4), 0.00005d, "少回收电能量电费", outputPath);
            VerifyGeneratedNumber(
                worksheet.Cell(row, 15),
                Math.Round(expected.TaxRate, 10),
                ChongqingStage2ExcelUtil.TaxRateTolerance,
                "扣税率",
                outputPath);
            VerifyGeneratedFormula(worksheet.Cell(row, 14), "K" + row + "-M" + row, "实际收益", outputPath);
            VerifyGeneratedNumber(worksheet.Cell(row, 14), Math.Round(expected.CalculatedNet, 4), Stage2SettlementCalculator.AmountTolerance, "实际收益", outputPath);
        }

        private static void VerifyGeneratedFormula(
            IXLCell cell,
            string expectedFormula,
            string label,
            string outputPath)
        {
            if (!string.Equals(TextUtil.S(cell.FormulaA1), expectedFormula, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidDataException(
                    "重庆阶段二分表" + label + "公式不符合预期：" + outputPath
                    + "，" + cell.Address.ToStringRelative() + "。");
            }
        }

        private static void VerifyGeneratedNumber(
            IXLCell cell,
            double expected,
            double tolerance,
            string label,
            string outputPath)
        {
            var actual = ClosedXmlUtil.CellNumber(cell);
            if (Math.Abs(actual - expected) > tolerance)
            {
                throw new InvalidDataException(
                    "重庆阶段二分表" + label + "与台账计划不一致：" + outputPath
                    + "，" + cell.Address.ToStringRelative() + "。"
                    + "预期=" + expected.ToString("0.######", CultureInfo.InvariantCulture)
                    + "，实际=" + actual.ToString("0.######", CultureInfo.InvariantCulture));
            }
        }

        public static List<GroupSettlementTotal> BuildSplitFiles(
            ChongqingStage2Options options,
            IList<ChongqingSettlementDetail> details,
            IList<string> warnings)
        {
            var templateMap = BuildTemplateIndex(options, warnings);
            var groups = new List<GroupSettlementTotal>();
            foreach (var group in details
                .GroupBy(detail => ChongqingStage2Keys.SummaryKey(detail.Entity, detail.Kind))
                .OrderBy(group => group.First().Kind)
                .ThenBy(group => group.First().Owner)
                .ThenBy(group => group.First().Entity))
            {
                var first = group.First();
                bool matchedTemplate;
                var outputPath = EnsureOutputWorkbook(templateMap, options, first.Kind, first.Owner, first.Entity, warnings, out matchedTemplate);
                FileAccessGuard.RequireWritableWorkbook(outputPath, ChongqingStage2ExcelUtil.KindShort(first.Kind) + "分表输出文件");

                using (var workbook = new XLWorkbook(outputPath))
                {
                    var worksheet = PrepareMonthSheet(workbook, options.Month);
                    WriteSplitSheet(worksheet, first.Kind, first.Entity, options.Month, group.ToList(), outputPath, warnings);
                    if (!matchedTemplate)
                    {
                        KeepOnlyCurrentMonthSheet(workbook, worksheet);
                    }

                    ClosedXmlUtil.SetOnlyActiveWorksheet(workbook, worksheet);
                    ChongqingStage2ExcelUtil.SaveWorkbook(workbook, outputPath);
                }

                groups.Add(new GroupSettlementTotal
                {
                    Kind = first.Kind,
                    Owner = first.Owner,
                    Entity = first.Entity,
                    DisplayEntity = first.Entity,
                    Rows = group.Count(),
                    ExpectedNet = Math.Round(group.Sum(item => item.ExpectedNet), 4),
                    OutputFile = outputPath
                });
            }

            return groups;
        }

        public static void AddTemplateIssues(
            ChongqingStage2Options options,
            IList<GroupSettlementTotal> groups,
            IList<ChongqingSettlementDetail> details,
            IList<ChongqingStage2CheckIssue> issues)
        {
            var catalog = ReadTemplateCatalog(options);
            foreach (var catalogIssue in catalog.Issues)
            {
                issues.Add(catalogIssue);
            }
            foreach (var group in groups)
            {
                var exact = catalog.Candidates
                    .Where(candidate => candidate.Kind == group.Kind
                        && TextUtil.CustomerKey(candidate.Entity) == TextUtil.CustomerKey(group.Entity))
                    .OrderBy(candidate => candidate.Path, StringComparer.OrdinalIgnoreCase)
                    .ToList();
                if (exact.Count > 1)
                {
                    issues.Add(new ChongqingStage2CheckIssue
                    {
                        Code = Stage2PreflightIssueKinds.DuplicateExactTemplates,
                        Disposition = Stage2PreflightDisposition.Blocker,
                        Severity = "阻断",
                        Category = "同一主体存在多个精确模板",
                        Kind = Stage2PreflightIssueKinds.DuplicateExactTemplates,
                        SettlementKind = group.Kind,
                        Owner = group.Owner,
                        Entity = group.Entity,
                        TemplateFile = string.Join("；", exact.Select(candidate => candidate.Path)),
                        Message = "同一费用类型和主体匹配到多个历史分表，程序无法可靠选择。",
                        Suggestion = "请只保留一个本次应继承的精确模板后重新预检。"
                    });
                    continue;
                }

                if (exact.Count == 1)
                {
                    var currentRows = details.Where(detail =>
                            ChongqingStage2Keys.SummaryKey(detail.Entity, detail.Kind)
                                == ChongqingStage2Keys.SummaryKey(group.Entity, group.Kind))
                        .ToList();
                    CompareGroupWithTemplate(
                        options.Month,
                        group,
                        exact[0].Path,
                        currentRows,
                        issues);
                    AddRefundExtraPowerRowIssue(
                        options.Month,
                        group,
                        exact[0].Path,
                        currentRows,
                        issues);
                    continue;
                }

                var borrowCandidates = catalog.Candidates
                    .Where(candidate => candidate.Kind == group.Kind)
                    .OrderBy(candidate => candidate.Path, StringComparer.OrdinalIgnoreCase)
                    .ToList();
                if (borrowCandidates.Count == 0)
                {
                    issues.Add(new ChongqingStage2CheckIssue
                    {
                        Code = Stage2PreflightIssueKinds.TemplateMissing,
                        Disposition = Stage2PreflightDisposition.Blocker,
                        Severity = "阻断",
                        Category = "没有可用的同类分表模板",
                        Kind = Stage2PreflightIssueKinds.TemplateMissing,
                        SettlementKind = group.Kind,
                        Owner = group.Owner,
                        Entity = group.Entity,
                        Message = "未找到该主体的精确模板，也没有可借用的同费用类型模板。",
                        Suggestion = "请补充可靠的" + group.Kind + "分表模板后重新预检。"
                    });
                    continue;
                }

                if (borrowCandidates.Count > 1)
                {
                    issues.Add(new ChongqingStage2CheckIssue
                    {
                        Code = Stage2PreflightIssueKinds.AmbiguousBorrowTemplates,
                        Disposition = Stage2PreflightDisposition.Blocker,
                        Severity = "阻断",
                        Category = "同类借用模板不唯一",
                        Kind = Stage2PreflightIssueKinds.AmbiguousBorrowTemplates,
                        SettlementKind = group.Kind,
                        Owner = group.Owner,
                        Entity = group.Entity,
                        TemplateFile = string.Join("；", borrowCandidates.Select(candidate => candidate.Path)),
                        Message = "未找到该主体的精确模板，但同费用类型有多个可借用候选，程序无法可靠选择。",
                        Suggestion = "请为新主体准备唯一精确模板，或将借用范围整理为唯一可靠候选。"
                    });
                    continue;
                }

                var borrowed = borrowCandidates[0];

                issues.Add(new ChongqingStage2CheckIssue
                {
                    Code = Stage2PreflightIssueKinds.BorrowedTemplate,
                    Disposition = Stage2PreflightDisposition.Review,
                    Severity = "复核",
                    Category = "新主体借用同类模板",
                    Kind = Stage2PreflightIssueKinds.BorrowedTemplate,
                    SettlementKind = group.Kind,
                    Owner = group.Owner,
                    Entity = group.Entity,
                    TemplateFile = borrowed.Path,
                    Message = "未找到该主体的精确模板，程序将借用上述唯一同类模板。",
                    Suggestion = "生成后请复核主体抬头、签章和付款信息。"
                });

                AddRefundExtraPowerRowIssue(
                    options.Month,
                    group,
                    borrowed.Path,
                    details.Where(detail =>
                            ChongqingStage2Keys.SummaryKey(detail.Entity, detail.Kind)
                                == ChongqingStage2Keys.SummaryKey(group.Entity, group.Kind))
                        .ToList(),
                    issues);
            }
        }

        public static void AddManagedOutputIssues(
            ChongqingStage2Options options,
            IList<ChongqingManagedOutputPlanItem> planned,
            IList<ChongqingStage2CheckIssue> issues)
        {
            foreach (var collision in planned
                .GroupBy(item => Path.GetFullPath(item.Path), StringComparer.OrdinalIgnoreCase)
                .Where(group => group.Count() > 1))
            {
                var identities = collision
                    .Select(item => item.Kind + "主体“" + item.Entity + "”（负责人：“" + item.Owner + "”）")
                    .ToList();
                issues.Add(new ChongqingStage2CheckIssue
                {
                    Code = Stage2PreflightIssueKinds.PlannedOutputPathConflict,
                    Disposition = Stage2PreflightDisposition.Blocker,
                    Severity = "阻断",
                    Category = "本批分表输出路径冲突",
                    Kind = Stage2PreflightIssueKinds.PlannedOutputPathConflict,
                    TemplateFile = collision.Key,
                    CurrentValue = string.Join("；", identities),
                    Message = "多个不同汇总身份规划到同一个正式分表路径，程序无法保证互不覆盖："
                        + string.Join("；", identities) + "。",
                    Suggestion = "请调整主体名称、负责人或模板文件名，使每个汇总身份对应唯一正式工作簿后重新预检。"
                });
            }

            var findings = Stage2ManagedOutputInspector.InspectUnexpectedWorkbooks(
                new[]
                {
                    ChongqingStage2ExcelUtil.OutputRootFor(options, ChongqingStage2SettlementKinds.Proxy),
                    ChongqingStage2ExcelUtil.OutputRootFor(options, ChongqingStage2SettlementKinds.Intermediary),
                    ChongqingStage2ExcelUtil.OutputRootFor(options, ChongqingStage2SettlementKinds.Refund)
                },
                planned.Select(item => item.Path),
                ChongqingStage2ExcelUtil.MonthSheetName(options.Month));
            foreach (var finding in findings)
            {
                var isReview = finding.IsPlannedTargetMonth;
                issues.Add(new ChongqingStage2CheckIssue
                {
                    Code = finding.IsUnreadable
                        ? Stage2PreflightIssueKinds.ManagedOutputUnreadable
                        : isReview
                            ? Stage2PreflightIssueKinds.PlannedTargetMonthWorkbook
                            : Stage2PreflightIssueKinds.UnexpectedTargetMonthWorkbook,
                    Disposition = isReview
                        ? Stage2PreflightDisposition.Review
                        : Stage2PreflightDisposition.Blocker,
                    Severity = isReview ? "复核" : "阻断",
                    Category = isReview
                        ? "本批计划分表已含目标月份"
                        : finding.IsUnreadable
                            ? "受管分表输出无法检查"
                            : "非本批分表仍含目标月份",
                    Kind = finding.IsUnreadable
                        ? Stage2PreflightIssueKinds.ManagedOutputUnreadable
                        : isReview
                            ? Stage2PreflightIssueKinds.PlannedTargetMonthWorkbook
                            : Stage2PreflightIssueKinds.UnexpectedTargetMonthWorkbook,
                    TemplateFile = finding.Path,
                    SheetName = ChongqingStage2ExcelUtil.MonthSheetName(options.Month),
                    Message = finding.Message,
                    Suggestion = isReview
                        ? "请确认其中的人工修改已经回填到本次输入模板；继续后将由整包事务覆盖。"
                        : finding.IsUnreadable
                            ? "请修复、关闭占用或移出无法检查的工作簿/目录后重新预检。"
                            : "请人工确认旧路径，将其移出正式受管目录或删除其中本月工作表后重新预检；程序不会自动删除年度历史工作簿。"
                });
            }
        }

        public static List<ChongqingManagedOutputPlanItem> BuildManagedOutputPlan(
            ChongqingStage2Options options,
            IEnumerable<GroupSettlementTotal> groups)
        {
            var catalog = ReadTemplateCatalog(options);
            var result = new List<ChongqingManagedOutputPlanItem>();
            foreach (var group in groups)
            {
                var exact = catalog.Candidates
                    .Where(candidate => candidate.Kind == group.Kind
                        && TextUtil.CustomerKey(candidate.Entity) == TextUtil.CustomerKey(group.Entity))
                    .ToList();
                string fileName;
                if (exact.Count == 1)
                {
                    fileName = Path.GetFileName(exact[0].Path);
                }
                else if (exact.Count == 0)
                {
                    var borrowed = catalog.Candidates
                        .Where(candidate => candidate.Kind == group.Kind)
                        .ToList();
                    if (borrowed.Count != 1)
                    {
                        continue;
                    }

                    fileName = TextUtil.SafeFileName(group.Entity) + " 2026重庆.xlsx";
                }
                else
                {
                    continue;
                }

                result.Add(new ChongqingManagedOutputPlanItem
                {
                    Kind = group.Kind,
                    Entity = group.Entity,
                    Owner = group.Owner,
                    Path = Path.Combine(
                        ChongqingStage2ExcelUtil.OutputRootFor(options, group.Kind),
                        TextUtil.SafeFileName(group.Owner),
                        fileName)
                });
            }

            return result;
        }

        public static void EnsureManagedOutputStillSafe(
            ChongqingStage2Options options,
            IEnumerable<ChongqingManagedOutputPlanItem> planned)
        {
            var blockingFindings = Stage2ManagedOutputInspector.InspectUnexpectedWorkbooks(
                    new[]
                    {
                        ChongqingStage2ExcelUtil.OutputRootFor(options, ChongqingStage2SettlementKinds.Proxy),
                        ChongqingStage2ExcelUtil.OutputRootFor(options, ChongqingStage2SettlementKinds.Intermediary),
                        ChongqingStage2ExcelUtil.OutputRootFor(options, ChongqingStage2SettlementKinds.Refund)
                    },
                    planned.Select(item => item.Path),
                    ChongqingStage2ExcelUtil.MonthSheetName(options.Month))
                .Where(finding => !finding.IsPlannedTargetMonth)
                .ToList();
            if (blockingFindings.Count > 0)
            {
                throw new InvalidDataException(
                    "重庆阶段二发布前复检发现正式分表目录已变化，本批不会发布："
                    + string.Join("；", blockingFindings.Select(finding => finding.Message).Take(10)));
            }
        }

        private static void AddRefundExtraPowerRowIssue(
            int month,
            GroupSettlementTotal group,
            string templatePath,
            IList<ChongqingSettlementDetail> currentRows,
            IList<ChongqingStage2CheckIssue> issues)
        {
            if (group.Kind != ChongqingStage2SettlementKinds.Refund)
            {
                return;
            }

            try
            {
                using (var workbook = ChongqingStage2ExcelUtil.OpenWorkbookShared(templatePath))
                {
                    var targetTitle = ChongqingStage2ExcelUtil.MonthSheetName(month);
                    var worksheet = workbook.Worksheets.FirstOrDefault(sheet => sheet.Name == targetTitle)
                        ?? PreviousMonthSheet(workbook, month, targetTitle)
                        ?? LastMonthSheet(workbook);
                    var firstTotalRow = FindSplitTotalRow(worksheet);
                    var extraRows = FindRefundExtraPowerRows(worksheet, firstTotalRow);
                    if (extraRows.Count == 0)
                    {
                        return;
                    }

                    var canSync = currentRows.Count == 1;
                    issues.Add(new ChongqingStage2CheckIssue
                    {
                        Code = ChongqingStage2IssueKinds.RefundExtraPowerRows,
                        Disposition = Stage2PreflightDisposition.Review,
                        Severity = "复核",
                        Category = "退补分表存在额外扣减块",
                        Kind = ChongqingStage2IssueKinds.RefundExtraPowerRows,
                        SettlementKind = group.Kind,
                        Owner = group.Owner,
                        Entity = group.Entity,
                        TemplateFile = templatePath,
                        SheetName = worksheet.Name,
                        PreviousValue = "额外扣减行：" + string.Join("、", extraRows.Select(row => row.ToString(CultureInfo.InvariantCulture))),
                        CurrentValue = "本月退补明细：" + currentRows.Count,
                        Message = canSync
                            ? "退补分表检测到额外扣减块；当前仅有 1 条退补明细，程序将把该明细的当月电量同步到额外块 C-G。"
                            : "退补分表检测到额外扣减块，但当前有 " + currentRows.Count
                                + " 条退补明细，程序无法可靠判断额外块应对应哪条明细，不会自动推导该块。",
                        Suggestion = canSync
                            ? "H 列以后以及汇总表的当月抵扣、实际支付仍按模板保留，请生成后人工复核。"
                            : "请生成后人工检查额外扣减块全部字段，以及汇总表当月抵扣和实际支付。"
                    });
                }
            }
            catch (Exception ex)
            {
                if (issues.Any(issue =>
                    issue.Code == Stage2PreflightIssueKinds.TemplateUnreadable
                    && string.Equals(issue.TemplateFile, templatePath, StringComparison.OrdinalIgnoreCase)))
                {
                    return;
                }

                issues.Add(new ChongqingStage2CheckIssue
                {
                    Code = Stage2PreflightIssueKinds.TemplateUnreadable,
                    Disposition = Stage2PreflightDisposition.Blocker,
                    Severity = "阻断",
                    Category = "退补附加块预检失败",
                    Kind = Stage2PreflightIssueKinds.TemplateUnreadable,
                    SettlementKind = group.Kind,
                    Owner = group.Owner,
                    Entity = group.Entity,
                    TemplateFile = templatePath,
                    Message = "读取本次实际选用的退补分表工作表失败：" + ex.Message,
                    Suggestion = "请修复该模板后重新预检。"
                });
            }
        }

        private static void CompareGroupWithTemplate(
            int month,
            GroupSettlementTotal group,
            string templatePath,
            IList<ChongqingSettlementDetail> currentRows,
            IList<ChongqingStage2CheckIssue> issues)
        {
            try
            {
                using (var workbook = ChongqingStage2ExcelUtil.OpenWorkbookShared(templatePath))
                {
                    var previousSheet = PreviousMonthSheet(
                        workbook,
                        month,
                        ChongqingStage2ExcelUtil.MonthSheetName(month))
                        ?? LastMonthSheet(workbook);
                    if (previousSheet == null)
                    {
                        throw new InvalidDataException("模板没有可用于纵向比较的月份工作表。");
                    }

                    var previousRows = ReadPreviousDetails(previousSheet, group.Kind);
                    var currentCustomers = new HashSet<string>(
                        currentRows.Select(row => TextUtil.CustomerKey(row.Customer)));
                    foreach (var current in currentRows)
                    {
                        PreviousSplitDetail previous;
                        if (!previousRows.TryGetValue(TextUtil.CustomerKey(current.Customer), out previous))
                        {
                            issues.Add(new ChongqingStage2CheckIssue
                            {
                                Code = Stage2PreflightIssueKinds.NewCustomer,
                                Disposition = Stage2PreflightDisposition.Review,
                                Severity = "复核",
                                Category = "客户本月新增到分表",
                                Kind = Stage2PreflightIssueKinds.NewCustomer,
                                SettlementKind = group.Kind,
                                Customer = current.Customer,
                                Owner = group.Owner,
                                Entity = group.Entity,
                                LedgerRow = current.LedgerRow,
                                TemplateFile = templatePath,
                                SheetName = previousSheet.Name,
                                Message = group.Kind + "主体“" + group.Entity + "”下的客户“"
                                    + current.Customer + "”在上月分表中未找到。",
                                Suggestion = "如果这是本月新增客户可以继续；否则请检查台账和上月分表中的客户名称。"
                            });
                            continue;
                        }

                        AddTemplateValueChangeIssue(issues, group, current, previous, templatePath, "电量比例", previous.Ratio, current.Ratio);
                        if (group.Kind == ChongqingStage2SettlementKinds.Refund)
                        {
                            AddTemplateValueChangeIssue(issues, group, current, previous, templatePath, "尖峰单价", previous.RefundSharpPrice, current.RefundSharpPrice);
                            AddTemplateValueChangeIssue(issues, group, current, previous, templatePath, "峰段单价", previous.RefundPeakPrice, current.RefundPeakPrice);
                            AddTemplateValueChangeIssue(issues, group, current, previous, templatePath, "平段单价", previous.RefundFlatPrice, current.RefundFlatPrice);
                            AddTemplateValueChangeIssue(issues, group, current, previous, templatePath, "谷段单价", previous.RefundValleyPrice, current.RefundValleyPrice);
                        }
                        else
                        {
                            AddTemplateValueChangeIssue(issues, group, current, previous, templatePath, "利润单价", previous.UnitPrice, current.UnitPrice);
                        }

                        AddTemplateValueChangeIssue(
                            issues,
                            group,
                            current,
                            previous,
                            templatePath,
                            "扣税率",
                            previous.TaxRate,
                            current.TaxRate,
                            ChongqingStage2ExcelUtil.TaxRateTolerance);
                    }

                    foreach (var previous in previousRows.Values.Where(item =>
                        !currentCustomers.Contains(TextUtil.CustomerKey(item.Customer))))
                    {
                        issues.Add(new ChongqingStage2CheckIssue
                        {
                            Code = Stage2PreflightIssueKinds.PreviousTemplateCustomerMissing,
                            Disposition = Stage2PreflightDisposition.Review,
                            Severity = "复核",
                            Category = "上月分表存在本月台账外明细行",
                            Kind = Stage2PreflightIssueKinds.PreviousTemplateCustomerMissing,
                            SettlementKind = group.Kind,
                            Customer = previous.Customer,
                            Owner = group.Owner,
                            Entity = group.Entity,
                            TemplateFile = templatePath,
                            SheetName = previous.SheetName,
                            PreviousValue = "上月分表第" + previous.Row + "行",
                            CurrentValue = "本月台账无匹配客户",
                            Message = group.Kind + "主体“" + group.Entity + "”的上月分表明细“"
                                + previous.Customer + "”未在本月台账中找到。",
                            Suggestion = "程序不会把该行继承到本月；如果仍需退补或补扣，请生成后人工同步分表和汇总表。"
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                issues.Add(new ChongqingStage2CheckIssue
                {
                    Code = Stage2PreflightIssueKinds.TemplateUnreadable,
                    Disposition = Stage2PreflightDisposition.Blocker,
                    Severity = "阻断",
                    Category = "上月分表纵向预检失败",
                    Kind = Stage2PreflightIssueKinds.TemplateUnreadable,
                    SettlementKind = group.Kind,
                    Owner = group.Owner,
                    Entity = group.Entity,
                    TemplateFile = templatePath,
                    Message = "读取上月分表客户和关系参数失败：" + ex.Message,
                    Suggestion = "请确认模板包含可读取的月份明细和合计行，修复后重新预检。"
                });
            }
        }

        private static Dictionary<string, PreviousSplitDetail> ReadPreviousDetails(
            IXLWorksheet worksheet,
            string kind)
        {
            var result = new Dictionary<string, PreviousSplitDetail>();
            var totalRow = FindSplitTotalRow(worksheet);
            for (var row = ChongqingStage2Layout.DetailDataStartRow; row < totalRow; row++)
            {
                var customer = ChongqingStage2ExcelUtil.CellText(worksheet.Cell(row, 2));
                if (string.IsNullOrWhiteSpace(customer))
                {
                    continue;
                }

                var key = TextUtil.CustomerKey(customer);
                if (result.ContainsKey(key))
                {
                    continue;
                }

                result[key] = new PreviousSplitDetail
                {
                    Row = row,
                    Customer = customer,
                    SheetName = worksheet.Name,
                    Ratio = ClosedXmlUtil.CellNumber(worksheet.Cell(row, 8)),
                    UnitPrice = kind == ChongqingStage2SettlementKinds.Refund
                        ? 0d
                        : ClosedXmlUtil.CellNumber(worksheet.Cell(row, 9)),
                    RefundSharpPrice = kind == ChongqingStage2SettlementKinds.Refund
                        ? ClosedXmlUtil.CellNumber(worksheet.Cell(row, 9))
                        : 0d,
                    RefundPeakPrice = kind == ChongqingStage2SettlementKinds.Refund
                        ? ClosedXmlUtil.CellNumber(worksheet.Cell(row, 10))
                        : 0d,
                    RefundFlatPrice = kind == ChongqingStage2SettlementKinds.Refund
                        ? ClosedXmlUtil.CellNumber(worksheet.Cell(row, 11))
                        : 0d,
                    RefundValleyPrice = kind == ChongqingStage2SettlementKinds.Refund
                        ? ClosedXmlUtil.CellNumber(worksheet.Cell(row, 12))
                        : 0d,
                    TaxRate = ClosedXmlUtil.CellNumber(
                        worksheet.Cell(row, kind == ChongqingStage2SettlementKinds.Refund ? 16 : 15))
                };
            }

            return result;
        }

        private static void AddTemplateValueChangeIssue(
            IList<ChongqingStage2CheckIssue> issues,
            GroupSettlementTotal group,
            ChongqingSettlementDetail current,
            PreviousSplitDetail previous,
            string templatePath,
            string fieldName,
            double previousValue,
            double currentValue,
            double tolerance = Stage2SettlementCalculator.AmountTolerance)
        {
            if (Math.Abs(previousValue - currentValue) <= tolerance)
            {
                return;
            }

            var previousDisplay = tolerance <= ChongqingStage2ExcelUtil.TaxRateTolerance
                ? previousValue.ToString("0.##########", CultureInfo.InvariantCulture)
                : Stage2SettlementCalculator.FormatAmount(previousValue);
            var currentDisplay = tolerance <= ChongqingStage2ExcelUtil.TaxRateTolerance
                ? currentValue.ToString("0.##########", CultureInfo.InvariantCulture)
                : Stage2SettlementCalculator.FormatAmount(currentValue);

            issues.Add(new ChongqingStage2CheckIssue
            {
                Code = Stage2PreflightIssueKinds.RelationshipValueChanged,
                Disposition = Stage2PreflightDisposition.Review,
                Severity = "复核",
                Category = "关键字段较上月变化",
                Kind = Stage2PreflightIssueKinds.RelationshipValueChanged,
                SettlementKind = group.Kind,
                Customer = current.Customer,
                Owner = group.Owner,
                Entity = group.Entity,
                LedgerRow = current.LedgerRow,
                TemplateFile = templatePath,
                SheetName = previous.SheetName,
                PreviousValue = fieldName + " 上月：" + previousDisplay,
                CurrentValue = fieldName + " 本月：" + currentDisplay,
                Message = group.Kind + "主体“" + group.Entity + "”下客户“" + current.Customer
                    + "”的" + fieldName + "发生变化（上月分表第" + previous.Row
                    + "行，本月台账第" + current.LedgerRow + "行）。",
                Suggestion = "如果这是本月正常调整可以继续；否则请回到台账检查该字段。"
            });
        }

        private static Dictionary<string, string> BuildTemplateIndex(ChongqingStage2Options options, IList<string> warnings)
        {
            var catalog = ReadTemplateCatalog(options);
            if (catalog.Issues.Count > 0)
            {
                throw new InvalidDataException("重庆阶段二分表模板在生成前发生变化或无法读取。");
            }

            var result = new Dictionary<string, string>();
            foreach (var duplicate in catalog.Candidates
                .GroupBy(candidate => ChongqingStage2Keys.SummaryKey(candidate.Entity, candidate.Kind))
                .Where(group => group.Count() > 1))
            {
                throw new InvalidDataException("重庆阶段二同一主体存在多个精确模板："
                    + string.Join("；", duplicate.Select(candidate => candidate.Path)));
            }

            foreach (var candidate in catalog.Candidates)
            {
                result.Add(
                    ChongqingStage2Keys.TemplateKey(candidate.Kind, candidate.Owner, candidate.Entity),
                    candidate.Path);
                result.Add(
                    ChongqingStage2Keys.TemplateKey(candidate.Kind, string.Empty, candidate.Entity),
                    candidate.Path);
            }

            return result;
        }

        private static ChongqingTemplateCatalog ReadTemplateCatalog(ChongqingStage2Options options)
        {
            var catalog = new ChongqingTemplateCatalog();
            IndexTemplateRoot(catalog, ChongqingStage2SettlementKinds.Proxy, options.ProxyTemplateDirectory, options.Month);
            IndexTemplateRoot(catalog, ChongqingStage2SettlementKinds.Intermediary, options.IntermediaryTemplateDirectory, options.Month);
            IndexTemplateRoot(catalog, ChongqingStage2SettlementKinds.Refund, options.RefundTemplateDirectory, options.Month);
            return catalog;
        }

        private static void IndexTemplateRoot(
            ChongqingTemplateCatalog catalog,
            string kind,
            string root,
            int month)
        {
            if (string.IsNullOrWhiteSpace(root) || !Directory.Exists(root))
            {
                return;
            }

            foreach (var path in Directory.GetFiles(root, "*.xlsx", SearchOption.AllDirectories)
                .OrderBy(item => item, StringComparer.OrdinalIgnoreCase))
            {
                var name = Path.GetFileName(path);
                if (name.StartsWith("~$", StringComparison.Ordinal) || name.StartsWith("._", StringComparison.Ordinal))
                {
                    continue;
                }

                try
                {
                    using (var workbook = ChongqingStage2ExcelUtil.OpenWorkbookShared(path))
                    {
                        var entity = ExtractEntityFromWorkbook(workbook);
                        if (string.IsNullOrWhiteSpace(entity))
                        {
                            throw new InvalidDataException("模板未识别到有效主体抬头。");
                        }

                        var comparisonSheet = PreviousMonthSheet(
                            workbook,
                            month,
                            ChongqingStage2ExcelUtil.MonthSheetName(month))
                            ?? LastMonthSheet(workbook);
                        if (comparisonSheet == null)
                        {
                            throw new InvalidDataException("模板没有可用的月份工作表。");
                        }

                        FindSplitTotalRow(comparisonSheet);

                        var owner = new DirectoryInfo(Path.GetDirectoryName(path)).Name;
                        catalog.Candidates.Add(new ChongqingTemplateCandidate
                        {
                            Kind = kind,
                            Entity = entity,
                            Owner = owner,
                            Path = path
                        });
                    }
                }
                catch (Exception ex)
                {
                    catalog.Issues.Add(new ChongqingStage2CheckIssue
                    {
                        Code = Stage2PreflightIssueKinds.TemplateUnreadable,
                        Disposition = Stage2PreflightDisposition.Blocker,
                        Severity = "阻断",
                        Category = "上月分表模板无法读取",
                        Kind = Stage2PreflightIssueKinds.TemplateUnreadable,
                        SettlementKind = kind,
                        TemplateFile = path,
                        Message = "读取重庆" + ChongqingStage2ExcelUtil.KindShort(kind) + "分表模板失败：" + ex.Message,
                        Suggestion = "请修复或移出损坏、无关的工作簿后重新预检。"
                    });
                }
            }
        }

        private static string ExtractEntityFromWorkbook(XLWorkbook workbook)
        {
            foreach (var sheet in workbook.Worksheets
                .Select(sheet =>
                {
                    int sheetMonth;
                    return new { Sheet = sheet, Matched = TryParseMonthSheet(sheet.Name, out sheetMonth), Month = sheetMonth };
                })
                .OrderByDescending(item => item.Matched)
                .ThenByDescending(item => item.Month)
                .Select(item => item.Sheet))
            {
                var entity = ExtractEntityFromSplitSheet(sheet);
                if (!string.IsNullOrWhiteSpace(entity))
                {
                    return entity;
                }
            }

            return string.Empty;
        }

        private static string EnsureOutputWorkbook(
            IDictionary<string, string> templateMap,
            ChongqingStage2Options options,
            string kind,
            string owner,
            string entity,
            IList<string> warnings,
            out bool matchedTemplate)
        {
            string source;
            if (templateMap.TryGetValue(ChongqingStage2Keys.TemplateKey(kind, owner, entity), out source)
                || templateMap.TryGetValue(ChongqingStage2Keys.TemplateKey(kind, string.Empty, entity), out source))
            {
                var targetRoot = ChongqingStage2ExcelUtil.OutputRootFor(options, kind);
                var target = Path.Combine(
                    targetRoot,
                    TextUtil.SafeFileName(owner),
                    Path.GetFileName(source));
                Directory.CreateDirectory(Path.GetDirectoryName(target));
                if (!File.Exists(target))
                {
                    ChongqingStage2ExcelUtil.CopyWorkbookShared(source, target, overwrite: false);
                }

                matchedTemplate = true;
                return target;
            }

            var borrowSources = templateMap
                .Where(item => item.Key.StartsWith(kind + "|", StringComparison.Ordinal))
                .Select(item => item.Value)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(item => item, StringComparer.OrdinalIgnoreCase)
                .ToList();
            if (borrowSources.Count == 0)
            {
                throw new InvalidOperationException("没有可用的重庆" + ChongqingStage2ExcelUtil.KindShort(kind) + "分表模板。");
            }

            if (borrowSources.Count > 1)
            {
                throw new InvalidOperationException("重庆" + ChongqingStage2ExcelUtil.KindShort(kind) + "借用模板不唯一，已停止生成。");
            }

            source = borrowSources[0];

            var folder = Path.Combine(ChongqingStage2ExcelUtil.OutputRootFor(options, kind), TextUtil.SafeFileName(owner));
            Directory.CreateDirectory(folder);
            var targetPath = Path.Combine(folder, TextUtil.SafeFileName(entity) + " 2026重庆.xlsx");
            if (!File.Exists(targetPath))
            {
                ChongqingStage2ExcelUtil.CopyWorkbookShared(source, targetPath, overwrite: false);
            }

            matchedTemplate = false;
            warnings.Add("新增" + kind + "分表主体：" + entity + "（负责人：" + owner + "），已借用同类模板生成，请人工复核抬头、签章和付款信息。");
            return targetPath;
        }

        private static void WriteSplitSheet(
            IXLWorksheet worksheet,
            string kind,
            string entity,
            int month,
            IList<ChongqingSettlementDetail> rows,
            string outputPath,
            IList<string> warnings)
        {
            SetSplitTopTitles(worksheet, kind, entity, month);
            var totalRow = AdjustDetailRows(worksheet, rows.Count);
            for (var index = 0; index < rows.Count; index++)
            {
                var excelRow = ChongqingStage2Layout.DetailDataStartRow + index;
                if (kind == ChongqingStage2SettlementKinds.Refund)
                {
                    WriteRefundDetailRow(worksheet, excelRow, index + 1, rows[index]);
                }
                else
                {
                    WriteProxyLikeDetailRow(worksheet, excelRow, index + 1, rows[index]);
                }

            }

            WriteSplitTotalRow(worksheet, totalRow, kind, rows.Count);
            if (kind == ChongqingStage2SettlementKinds.Refund)
            {
                SyncRefundExtraPowerRows(worksheet, rows, outputPath, warnings);
            }
        }

        private static void WriteProxyLikeDetailRow(IXLWorksheet worksheet, int row, int sequence, ChongqingSettlementDetail detail)
        {
            worksheet.Cell(row, 1).Value = sequence;
            worksheet.Cell(row, 2).Value = detail.Customer;
            worksheet.Cell(row, 3).Value = Math.Round(detail.Total, 4);
            worksheet.Cell(row, 4).Value = Math.Round(detail.Sharp, 4);
            worksheet.Cell(row, 5).Value = Math.Round(detail.Peak, 4);
            worksheet.Cell(row, 6).Value = Math.Round(detail.Flat, 4);
            worksheet.Cell(row, 7).Value = Math.Round(detail.Valley, 4);
            worksheet.Cell(row, 8).Value = Math.Round(detail.Ratio, 6);
            worksheet.Cell(row, 9).Value = Math.Round(detail.UnitPrice, 6);
            worksheet.Cell(row, 10).FormulaA1 = "ROUND(C" + row + "*H" + row + "*I" + row + "/10,4)";
            worksheet.Cell(row, 11).FormulaA1 = "J" + row + "-L" + row;
            worksheet.Cell(row, 12).Value = Math.Round(detail.RecoverShortfall, 4);
            worksheet.Cell(row, 13).FormulaA1 = "ROUND(K" + row + "/1.13*O" + row + ",4)";
            worksheet.Cell(row, 14).FormulaA1 = "K" + row + "-M" + row;
            worksheet.Cell(row, 15).Value = Math.Round(detail.TaxRate, 10);
        }

        private static void WriteRefundDetailRow(IXLWorksheet worksheet, int row, int sequence, ChongqingSettlementDetail detail)
        {
            worksheet.Cell(row, 1).Value = sequence;
            worksheet.Cell(row, 2).Value = detail.Customer;
            worksheet.Cell(row, 3).FormulaA1 = "SUM(D" + row + ":G" + row + ")";
            worksheet.Cell(row, 4).Value = Math.Round(detail.Sharp, 4);
            worksheet.Cell(row, 5).Value = Math.Round(detail.Peak, 4);
            worksheet.Cell(row, 6).Value = Math.Round(detail.Flat, 4);
            worksheet.Cell(row, 7).Value = Math.Round(detail.Valley, 4);
            worksheet.Cell(row, 8).Value = Math.Round(detail.Ratio, 6);
            worksheet.Cell(row, 9).Value = Math.Round(detail.RefundSharpPrice, 6);
            worksheet.Cell(row, 10).Value = Math.Round(detail.RefundPeakPrice, 6);
            worksheet.Cell(row, 11).Value = Math.Round(detail.RefundFlatPrice, 6);
            worksheet.Cell(row, 12).Value = Math.Round(detail.RefundValleyPrice, 6);
            worksheet.Cell(row, 13).FormulaA1 = "ROUND((D" + row + "*H" + row + "*I" + row + "+E" + row + "*H" + row + "*J" + row + "+F" + row + "*H" + row + "*K" + row + "+G" + row + "*H" + row + "*L" + row + ")/10,4)";
            worksheet.Cell(row, 14).FormulaA1 = "ROUND(M" + row + "/1.13*P" + row + ",4)";
            worksheet.Cell(row, 15).FormulaA1 = "M" + row + "-N" + row;
            worksheet.Cell(row, 16).Value = Math.Round(detail.TaxRate, 10);
        }

        private static void WriteSplitTotalRow(IXLWorksheet worksheet, int totalRow, string kind, int rowCount)
        {
            worksheet.Cell(totalRow, 1).Value = "合计";
            worksheet.Cell(totalRow, 2).Clear(XLClearOptions.Contents);
            if (rowCount == 0)
            {
                return;
            }

            var last = ChongqingStage2Layout.DetailDataStartRow + rowCount - 1;
            var sumColumns = kind == ChongqingStage2SettlementKinds.Refund
                ? new[] { 3, 4, 5, 6, 7, 13, 14, 15 }
                : new[] { 3, 4, 5, 6, 7, 10, 11, 12, 13, 14 };
            foreach (var column in sumColumns)
            {
                var letter = ClosedXmlUtil.ColumnLetter(column);
                worksheet.Cell(totalRow, column).FormulaA1 = "SUM(" + letter + ChongqingStage2Layout.DetailDataStartRow + ":" + letter + last + ")";
            }
        }

        private static IXLWorksheet PrepareMonthSheet(XLWorkbook workbook, int month)
        {
            var title = ChongqingStage2ExcelUtil.MonthSheetName(month);
            var existing = workbook.Worksheets.FirstOrDefault(sheet => sheet.Name == title);
            if (existing != null)
            {
                return existing;
            }

            var source = PreviousMonthSheet(workbook, month, title);
            if (source == null)
            {
                source = LastMonthSheet(workbook);
            }

            return source.CopyTo(title);
        }

        private static IXLWorksheet LastMonthSheet(XLWorkbook workbook)
        {
            var monthSheets = workbook.Worksheets
                .Select(sheet =>
                {
                    int sheetMonth;
                    return new { Sheet = sheet, Matched = TryParseMonthSheet(sheet.Name, out sheetMonth), Month = sheetMonth };
                })
                .Where(item => item.Matched)
                .OrderBy(item => item.Month)
                .ToList();
            return monthSheets.Count == 0 ? workbook.Worksheets.Last() : monthSheets.Last().Sheet;
        }

        private static IXLWorksheet PreviousMonthSheet(XLWorkbook workbook, int month, string targetTitle)
        {
            return workbook.Worksheets
                .Select(sheet =>
                {
                    int sheetMonth;
                    return new { Sheet = sheet, Matched = TryParseMonthSheet(sheet.Name, out sheetMonth), Month = sheetMonth };
                })
                .Where(item => item.Matched && item.Month < month && item.Sheet.Name != targetTitle)
                .OrderBy(item => item.Month)
                .Select(item => item.Sheet)
                .LastOrDefault();
        }

        private static bool TryParseMonthSheet(string name, out int month)
        {
            month = 0;
            var match = Regex.Match(TextUtil.S(name), "^(\\d{1,2})(?:\\s*\\(\\d+\\))?$");
            return match.Success && int.TryParse(match.Groups[1].Value, out month);
        }

        private static void KeepOnlyCurrentMonthSheet(XLWorkbook workbook, IXLWorksheet currentMonthSheet)
        {
            var currentName = currentMonthSheet.Name;
            foreach (var sheet in workbook.Worksheets.Where(sheet => sheet.Name != currentName).ToList())
            {
                sheet.Delete();
            }
        }

        private static int AdjustDetailRows(IXLWorksheet worksheet, int count)
        {
            var totalRow = FindSplitTotalRow(worksheet);
            var existing = totalRow - ChongqingStage2Layout.DetailDataStartRow;
            if (count > existing)
            {
                worksheet.Row(totalRow).InsertRowsAbove(count - existing);
                for (var row = totalRow; row < totalRow + count - existing; row++)
                {
                    worksheet.Row(ChongqingStage2Layout.DetailDataStartRow).CopyTo(worksheet.Row(row));
                }
            }
            else if (count < existing)
            {
                worksheet.Rows(ChongqingStage2Layout.DetailDataStartRow + count, totalRow - 1).Delete();
            }

            return FindSplitTotalRow(worksheet);
        }

        private static int FindSplitTotalRow(IXLWorksheet worksheet)
        {
            var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? ChongqingStage2Layout.DetailDataStartRow;
            for (var row = ChongqingStage2Layout.DetailDataStartRow; row <= lastRow; row++)
            {
                if (ChongqingStage2ExcelUtil.CellText(worksheet.Cell(row, 1)) == "合计")
                {
                    return row;
                }
            }

            for (var row = ChongqingStage2Layout.DetailDataStartRow; row <= lastRow; row++)
            {
                if (!string.IsNullOrWhiteSpace(ChongqingStage2ExcelUtil.CellText(worksheet.Cell(row, 2))))
                {
                    continue;
                }

                for (var column = 3; column <= Math.Min(16, worksheet.LastColumnUsed()?.ColumnNumber() ?? 16); column++)
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

        private static void SyncRefundExtraPowerRows(
            IXLWorksheet worksheet,
            IList<ChongqingSettlementDetail> rows,
            string outputPath,
            IList<string> warnings)
        {
            var firstTotalRow = FindSplitTotalRow(worksheet);
            var extraRows = FindRefundExtraPowerRows(worksheet, firstTotalRow);
            if (extraRows.Count == 0)
            {
                return;
            }

            var entity = rows.Count > 0 ? rows[0].Entity : ExtractEntityFromSplitSheet(worksheet);
            if (rows.Count != 1)
            {
                warnings.Add("重庆退补分表“" + entity + "”" + worksheet.Name + "月检测到" + extraRows.Count + "行额外扣减块，但当前主体有" + rows.Count + "条退补明细，无法安全自动匹配，已保留额外块 C-G 原值；请人工处理。文件：" + outputPath);
                return;
            }

            var detail = rows[0];
            foreach (var row in extraRows)
            {
                worksheet.Cell(row, 3).FormulaA1 = "SUM(D" + row + ":G" + row + ")";
                worksheet.Cell(row, 4).Value = Math.Round(detail.Sharp, 4);
                worksheet.Cell(row, 5).Value = Math.Round(detail.Peak, 4);
                worksheet.Cell(row, 6).Value = Math.Round(detail.Flat, 4);
                worksheet.Cell(row, 7).Value = Math.Round(detail.Valley, 4);
            }

            warnings.Add("重庆退补分表“" + entity + "”" + worksheet.Name + "月检测到" + extraRows.Count + "行额外扣减块，已同步 C-G 当月电量；H列以后、汇总表当月抵扣和实际支付仍按模板保留，请人工复核。文件：" + outputPath);
        }

        private static List<int> FindRefundExtraPowerRows(IXLWorksheet worksheet, int firstTotalRow)
        {
            var result = new List<int>();
            var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? firstTotalRow;
            for (var row = firstTotalRow + 1; row <= lastRow; row++)
            {
                var label = ChongqingStage2ExcelUtil.CellText(worksheet.Cell(row, 2));
                if (label.IndexOf("当月应扣", StringComparison.Ordinal) >= 0
                    || label.IndexOf("电表改造", StringComparison.Ordinal) >= 0)
                {
                    result.Add(row);
                }
            }

            return result;
        }

        private static void SetSplitTopTitles(IXLWorksheet worksheet, string kind, string entity, int month)
        {
            if (kind == ChongqingStage2SettlementKinds.Refund)
            {
                worksheet.Cell("A1").Value = "退补电费结算单";
                worksheet.Cell("A2").Value = "名称:" + entity;
            }
            else
            {
                worksheet.Cell("A1").Value = ChongqingStage2ExcelUtil.KindShort(kind) + "费用结算单";
                worksheet.Cell("A2").Value = "代理名称:" + entity;
            }

            SetFirstCellContaining(worksheet, "所属期", "所属期：" + ChongqingStage2Layout.Year + " 年 " + month.ToString("00", CultureInfo.InvariantCulture) + " 月");
            var nextMonth = month == 12 ? 1 : month + 1;
            var year = month == 12 ? ChongqingStage2Layout.Year + 1 : ChongqingStage2Layout.Year;
            SetFirstCellContaining(worksheet, "结算日期", "结算日期：" + year + " 年 " + nextMonth.ToString("00", CultureInfo.InvariantCulture) + " 月 15 日");
        }

        private static void SetFirstCellContaining(IXLWorksheet worksheet, string needle, string value)
        {
            var maxRow = Math.Min(worksheet.LastRowUsed()?.RowNumber() ?? 1, 8);
            var maxColumn = Math.Min(worksheet.LastColumnUsed()?.ColumnNumber() ?? 1, 24);
            for (var row = 1; row <= maxRow; row++)
            {
                for (var column = 1; column <= maxColumn; column++)
                {
                    if (ChongqingStage2ExcelUtil.CellText(worksheet.Cell(row, column)).Contains(needle))
                    {
                        worksheet.Cell(row, column).Value = value;
                        return;
                    }
                }
            }
        }

        private static string ExtractEntityFromSplitSheet(IXLWorksheet sheet)
        {
            var text = ChongqingStage2ExcelUtil.CellText(sheet.Cell("A2"));
            foreach (var prefix in new[] { "代理名称:", "代理名称：", "居间名称:", "居间名称：", "名称:", "名称：" })
            {
                if (text.StartsWith(prefix, StringComparison.Ordinal))
                {
                    return TextUtil.S(text.Substring(prefix.Length));
                }
            }

            return string.Empty;
        }

        private sealed class ChongqingTemplateCatalog
        {
            public List<ChongqingTemplateCandidate> Candidates { get; } =
                new List<ChongqingTemplateCandidate>();

            public List<ChongqingStage2CheckIssue> Issues { get; } =
                new List<ChongqingStage2CheckIssue>();
        }

        private sealed class ChongqingTemplateCandidate
        {
            public string Kind { get; set; }
            public string Entity { get; set; }
            public string Owner { get; set; }
            public string Path { get; set; }
        }

        private sealed class PreviousSplitDetail
        {
            public int Row { get; set; }
            public string Customer { get; set; }
            public string SheetName { get; set; }
            public double Ratio { get; set; }
            public double UnitPrice { get; set; }
            public double RefundSharpPrice { get; set; }
            public double RefundPeakPrice { get; set; }
            public double RefundFlatPrice { get; set; }
            public double RefundValleyPrice { get; set; }
            public double TaxRate { get; set; }
        }
    }

    internal sealed class ChongqingManagedOutputPlanItem
    {
        public string Kind { get; set; }
        public string Entity { get; set; }
        public string Owner { get; set; }
        public string Path { get; set; }
    }
}
