using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using HainanSettlementTool.Core.Models;
using HainanSettlementTool.Core.Services;

namespace HainanSettlementTool.Excel
{
    internal sealed class HainanStage2SettlementGenerator
    {
        public HainanStage2Report Generate(HainanStage2Options options)
        {
            Directory.CreateDirectory(options.OutputDirectory);
            var proxyRows = new List<HainanStage2DetailSettlementRow>();
            var interRows = new List<HainanStage2DetailSettlementRow>();
            HainanStage2LedgerReader.ReadLedgerRows(options.LedgerPath, options.Month, proxyRows, interRows);

            var missingOwners = proxyRows
                .Concat(interRows)
                .Where(row => string.IsNullOrWhiteSpace(row.Owner))
                .Select(row => "第" + row.LedgerRow + "行 " + row.Customer + " 缺少负责人")
                .Distinct()
                .ToList();
            if (missingOwners.Count > 0 && !options.AllowMissingOwner)
            {
                throw new InvalidOperationException(options.Month + "月结算明细存在负责人缺失，先不要生成分表/汇总表：" + string.Join("、", missingOwners.Take(10)));
            }

            var auditIssues = BuildPreflightIssues(options, proxyRows, interRows);
            ValidateRequiredPaymentPartyDecisions(options, auditIssues);
            var totals = HainanStage2SplitWorkbookWriter.BuildSplitFiles(options, proxyRows, interRows, auditIssues);
            var warnings = new List<string>();
            var summaryPath = HainanStage2SummaryWorkbookWriter.BuildSummary(options, totals, warnings);
            var report = HainanStage2ReportWriter.CreateReport(options, proxyRows, interRows, totals, summaryPath, warnings, missingOwners, auditIssues);
            HainanStage2ReportWriter.WriteReport(options, report);
            HainanStage2ReportWriter.WriteWarnings(options, warnings);
            HainanStage2ReportWriter.WriteAuditReport(options, report);
            return report;
        }

        public HainanStage2PreflightReport Analyze(HainanStage2Options options)
        {
            var proxyRows = new List<HainanStage2DetailSettlementRow>();
            var interRows = new List<HainanStage2DetailSettlementRow>();
            HainanStage2LedgerReader.ReadLedgerRows(options.LedgerPath, options.Month, proxyRows, interRows);

            var report = new HainanStage2PreflightReport
            {
                Month = options.Month
            };
            report.Issues.AddRange(BuildPreflightIssues(options, proxyRows, interRows));
            return report;
        }

        private static List<HainanStage2CheckIssue> BuildPreflightIssues(HainanStage2Options options, IList<HainanStage2DetailSettlementRow> proxyRows, IList<HainanStage2DetailSettlementRow> interRows)
        {
            var issues = new List<HainanStage2CheckIssue>();
            var templateMap = HainanStage2TemplateIndex.Build(options.ProxyTemplateDirectory, options.IntermediaryTemplateDirectory);
            var grouped = proxyRows
                .Select(row => new { Key = Tuple.Create("代理", row.Owner, row.Entity), Row = row })
                .Concat(interRows.Select(row => new { Key = Tuple.Create("居间", row.Owner, row.Entity), Row = row }))
                .GroupBy(item => item.Key)
                .OrderBy(group => group.Key.Item1)
                .ThenBy(group => group.Key.Item2)
                .ThenBy(group => group.Key.Item3);

            foreach (var group in grouped)
            {
                var kind = group.Key.Item1;
                var owner = group.Key.Item2;
                var entity = group.Key.Item3;
                var templateKey = HainanStage2ExcelUtil.TemplateKey(kind, owner, entity);
                string templatePath;
                if (!templateMap.TryGetValue(templateKey, out templatePath))
                {
                    issues.Add(new HainanStage2CheckIssue
                    {
                        Severity = "提示",
                        Category = "未匹配到上月分表模板",
                        Kind = kind + "费",
                        Owner = owner,
                        Entity = entity,
                        Message = kind + "费主体“" + entity + "”未在上月分表文件夹中匹配到同名模板。",
                        Suggestion = "程序会复制同类型模板生成新分表；请确认这是新增关系，或检查负责人文件夹、分表文件名、A2代理名称是否和台账一致。"
                    });
                    continue;
                }

                CompareGroupWithTemplate(options.Month, kind, owner, entity, templatePath, group.Select(item => item.Row).ToList(), issues);
            }

            AddSummarySubjectPaymentIssues(options, proxyRows, interRows, issues);
            return issues;
        }

        private static void AddSummarySubjectPaymentIssues(
            HainanStage2Options options,
            IList<HainanStage2DetailSettlementRow> proxyRows,
            IList<HainanStage2DetailSettlementRow> interRows,
            IList<HainanStage2CheckIssue> issues)
        {
            var subjects = BuildExpectedSummarySubjects(proxyRows, interRows);
            if (subjects.Count == 0)
            {
                return;
            }

            using (var workbook = new XLWorkbook(options.SummaryTemplatePath))
            {
                var mainSheetName = HainanStage2SummaryWorkbookWriter.ResolveSummarySheetName(workbook, "main", true);
                var knownKeys = new HashSet<string>(HainanStage2SummaryWorkbookWriter.ReadSummaryMeta(workbook.Worksheet(mainSheetName))
                    .Select(item => HainanStage2ExcelUtil.SummaryKey(item.Entity, item.Kind)));

                foreach (var subject in subjects)
                {
                    if (knownKeys.Contains(HainanStage2ExcelUtil.SummaryKey(subject.Entity, subject.Kind)))
                    {
                        continue;
                    }

                    string paymentParty;
                    if (HainanStage2ExcelUtil.TryGetPaymentPartyOverride(subject.Entity, subject.Kind, options.Month, out paymentParty)
                        || HainanStage2ExcelUtil.TryGetPaymentPartyDecision(options, subject.Entity, subject.Kind, out paymentParty))
                    {
                        continue;
                    }

                    var issue = new HainanStage2CheckIssue
                    {
                        Severity = "确认",
                        Category = "新增汇总主体支付方选择",
                        Kind = subject.Kind,
                        Owner = subject.Owner,
                        Entity = subject.Entity,
                        TemplateFile = options.SummaryTemplatePath,
                        SheetName = mainSheetName,
                        Message = "汇总表模板缺少" + subject.Kind + "主体“" + subject.Entity + "”，支付方无法从历史汇总主体继承。",
                        Suggestion = "请选择本月生成时写入的支付方；本次选择只用于输出汇总表副本。",
                        RequiresPaymentPartySelection = true
                    };
                    issue.AvailablePaymentParties.AddRange(HainanStage2PaymentParties.Supported);
                    issues.Add(issue);
                }
            }
        }

        private static List<SummarySubject> BuildExpectedSummarySubjects(
            IList<HainanStage2DetailSettlementRow> proxyRows,
            IList<HainanStage2DetailSettlementRow> interRows)
        {
            return proxyRows
                .Select(row => new SummarySubject { Kind = "代理费", Owner = row.Owner, Entity = row.Entity })
                .Concat(interRows.Select(row => new SummarySubject { Kind = "居间费", Owner = row.Owner, Entity = row.Entity }))
                .GroupBy(subject => HainanStage2ExcelUtil.SummaryKey(subject.Entity, subject.Kind))
                .Select(group => group.First())
                .OrderBy(subject => subject.Kind)
                .ThenBy(subject => subject.Owner)
                .ThenBy(subject => subject.Entity)
                .ToList();
        }

        private static void ValidateRequiredPaymentPartyDecisions(
            HainanStage2Options options,
            IList<HainanStage2CheckIssue> issues)
        {
            var missing = issues
                .Where(issue => issue.RequiresPaymentPartySelection)
                .Where(issue =>
                {
                    string paymentParty;
                    return !HainanStage2ExcelUtil.TryGetPaymentPartyDecision(options, issue.Entity, issue.Kind, out paymentParty);
                })
                .Select(issue => issue.Kind + " " + issue.Entity)
                .Distinct()
                .ToList();

            if (missing.Count > 0)
            {
                throw new InvalidOperationException("海南阶段二新增汇总主体支付方未选择：" + string.Join("、", missing) + "。请在预检窗口选择清能或清辉后再生成。");
            }
        }

        private static void CompareGroupWithTemplate(
            int month,
            string kind,
            string owner,
            string entity,
            string templatePath,
            IList<HainanStage2DetailSettlementRow> currentRows,
            IList<HainanStage2CheckIssue> issues)
        {
            try
            {
                using (var workbook = new XLWorkbook(templatePath))
                {
                    var previousSheet = HainanStage2ExcelUtil.PreviousMonthSheet(workbook, month, month + "月") ?? HainanStage2ExcelUtil.LastMonthSheet(workbook);
                    var previousRows = ReadPreviousDetails(previousSheet);
                    var currentCustomers = new HashSet<string>(currentRows.Select(row => HainanStage2ExcelUtil.NormalizeName(row.Customer)));
                    foreach (var row in currentRows)
                    {
                        PreviousDetailRow previous;
                        if (!previousRows.TryGetValue(HainanStage2ExcelUtil.NormalizeName(row.Customer), out previous))
                        {
                            issues.Add(new HainanStage2CheckIssue
                            {
                                Severity = "提示",
                                Category = "客户本月新增到分表",
                                Kind = kind + "费",
                                Customer = row.Customer,
                                Owner = owner,
                                Entity = entity,
                                LedgerRow = row.LedgerRow,
                                TemplateFile = templatePath,
                                SheetName = previousSheet.Name,
                                Message = kind + "费主体“" + entity + "”下的客户“" + row.Customer + "”在上月分表未找到。",
                                Suggestion = "如果这是本月新增客户，可以继续；否则请检查台账客户名称和上月分表客户名称是否一致。"
                            });
                            continue;
                        }

                        AddValueChangeIssue(issues, kind, owner, entity, row, previous, templatePath, "电量比例", previous.Ratio, row.Ratio);
                        AddValueChangeIssue(issues, kind, owner, entity, row, previous, templatePath, "利润单价", previous.UnitPrice, row.UnitPrice);
                        AddValueChangeIssue(issues, kind, owner, entity, row, previous, templatePath, "税率", previous.TaxRate, row.TaxRate);
                    }

                    foreach (var previous in previousRows.Values.Where(row => !currentCustomers.Contains(HainanStage2ExcelUtil.NormalizeName(row.Customer))))
                    {
                        issues.Add(new HainanStage2CheckIssue
                        {
                            Severity = "提示",
                            Category = "上月分表存在本月台账外明细行",
                            Kind = kind + "费",
                            Customer = previous.Customer,
                            Owner = owner,
                            Entity = entity,
                            TemplateFile = templatePath,
                            SheetName = previous.SheetName,
                            PreviousValue = "上月分表第" + previous.Row + "行",
                            CurrentValue = "本月台账无匹配客户",
                            Message = kind + "费主体“" + entity + "”的上月分表第" + previous.Row + "行“" + previous.Customer + "”未在本月台账中匹配到。",
                            Suggestion = "程序生成本月分表时不会继承该行；如果本月仍需退补或补扣，请生成后手动调整分表和汇总表。"
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                issues.Add(new HainanStage2CheckIssue
                {
                    Severity = "提示",
                    Category = "上月分表预检失败",
                    Kind = kind + "费",
                    Owner = owner,
                    Entity = entity,
                    TemplateFile = templatePath,
                    Message = "读取上月分表模板失败：" + ex.Message,
                    Suggestion = "程序仍可继续生成；请确认该分表文件没有损坏，且包含上月明细和合计行。"
                });
            }
        }

        private static Dictionary<string, PreviousDetailRow> ReadPreviousDetails(IXLWorksheet worksheet)
        {
            var result = new Dictionary<string, PreviousDetailRow>();
            var totalRow = HainanStage2ExcelUtil.FindTotalRow(worksheet, HainanStage2ExcelUtil.DataStartRow);
            for (var row = HainanStage2ExcelUtil.DataStartRow; row < totalRow; row++)
            {
                var customer = TextUtil.S(worksheet.Cell(row, 2).GetFormattedString());
                if (string.IsNullOrWhiteSpace(customer))
                {
                    continue;
                }

                var key = HainanStage2ExcelUtil.NormalizeName(customer);
                if (result.ContainsKey(key))
                {
                    continue;
                }

                result[key] = new PreviousDetailRow
                {
                    Customer = customer,
                    Row = row,
                    SheetName = worksheet.Name,
                    Ratio = HainanStage2ExcelUtil.GetNumeric(worksheet, row, 10),
                    UnitPrice = HainanStage2ExcelUtil.GetNumeric(worksheet, row, 11),
                    TaxRate = HainanStage2ExcelUtil.GetNumeric(worksheet, row, 17)
                };
            }

            return result;
        }

        private static void AddValueChangeIssue(
            IList<HainanStage2CheckIssue> issues,
            string kind,
            string owner,
            string entity,
            HainanStage2DetailSettlementRow current,
            PreviousDetailRow previous,
            string templatePath,
            string fieldName,
            double previousValue,
            double currentValue)
        {
            if (Math.Abs(previousValue - currentValue) <= Stage2SettlementCalculator.AmountTolerance)
            {
                return;
            }

            issues.Add(new HainanStage2CheckIssue
            {
                Severity = "确认",
                Category = "关键字段较上月变化",
                Kind = kind + "费",
                Customer = current.Customer,
                Owner = owner,
                Entity = entity,
                LedgerRow = current.LedgerRow,
                TemplateFile = templatePath,
                SheetName = previous.SheetName,
                PreviousValue = fieldName + " 上月：" + Stage2SettlementCalculator.FormatAmount(previousValue),
                CurrentValue = fieldName + " 本月：" + Stage2SettlementCalculator.FormatAmount(currentValue),
                Message = kind + "费主体“" + entity + "”下的客户“" + current.Customer + "”" + fieldName + "发生变化（上月分表第" + previous.Row + "行，本月台账第" + current.LedgerRow + "行）。",
                Suggestion = "如果这是本月台账更新后的正常变化，请继续；否则请回到台账检查该客户的" + fieldName + "。"
            });
        }

        private sealed class SummarySubject
        {
            public string Entity { get; set; }
            public string Kind { get; set; }
            public string Owner { get; set; }
        }

        private sealed class PreviousDetailRow
        {
            public int Row { get; set; }
            public string Customer { get; set; }
            public string SheetName { get; set; }
            public double Ratio { get; set; }
            public double UnitPrice { get; set; }
            public double TaxRate { get; set; }
        }
    }
}
