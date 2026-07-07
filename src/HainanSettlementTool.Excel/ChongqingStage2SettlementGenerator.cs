using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using HainanSettlementTool.Core.Models;
using HainanSettlementTool.Core.Services;

namespace HainanSettlementTool.Excel
{
    internal sealed class ChongqingStage2SettlementGenerator
    {
        private const int DataStartRow = 4;

        public ChongqingStage2PreflightReport Analyze(ChongqingStage2Options options)
        {
            var groups = ReadLedgerGroups(options);
            var report = new ChongqingStage2PreflightReport { Month = options.Month };
            AddSummaryPaymentIssues(options, groups, report.Issues);
            return report;
        }

        public ChongqingStage2Report Generate(ChongqingStage2Options options)
        {
            throw new NotImplementedException("重庆阶段二 Excel 生成器当前只实现台账读取和预检；分表、退补表和汇总表生成尚未实现。");
        }

        private static List<GroupSettlementTotal> ReadLedgerGroups(ChongqingStage2Options options)
        {
            using (var stream = File.Open(options.LedgerPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var workbook = new XLWorkbook(stream))
            {
                var worksheet = FindLedgerWorksheet(workbook);
                var map = FindLedgerMap(worksheet, options.Month);
                var details = ReadDetails(worksheet, map);
                return details
                    .GroupBy(detail => SummaryKey(detail.Entity, detail.Kind))
                    .Select(group =>
                    {
                        var first = group.First();
                        return new GroupSettlementTotal
                        {
                            Kind = first.Kind,
                            Owner = first.Owner,
                            Entity = first.Entity,
                            DisplayEntity = first.Entity,
                            Rows = group.Count(),
                            ExpectedNet = Math.Round(group.Sum(item => item.ExpectedNet), 4)
                        };
                    })
                    .OrderBy(group => group.Kind)
                    .ThenBy(group => group.Owner)
                    .ThenBy(group => group.Entity)
                    .ToList();
            }
        }

        private static List<ChongqingSettlementDetail> ReadDetails(IXLWorksheet worksheet, LedgerMap map)
        {
            var details = new List<ChongqingSettlementDetail>();
            var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? DataStartRow;
            for (var row = DataStartRow; row <= lastRow; row++)
            {
                var customer = CellText(worksheet.Cell(row, map.CustomerNameColumn));
                if (string.IsNullOrWhiteSpace(customer))
                {
                    continue;
                }

                var owner = CellText(worksheet.Cell(row, map.OwnerColumn));
                var proxyEntity = CellText(worksheet.Cell(row, map.ProjectDeveloperColumn));
                var agentOrSelf = CellText(worksheet.Cell(row, map.AgentOrSelfColumn));
                var intermediaryEntity = map.IntermediaryColumn > 0
                    ? CellText(worksheet.Cell(row, map.IntermediaryColumn))
                    : string.Empty;
                var refundEntity = map.PayeeColumn > 0
                    ? CellText(worksheet.Cell(row, map.PayeeColumn))
                    : string.Empty;

                AddDetailIfNeeded(
                    details,
                    row,
                    customer,
                    owner,
                    proxyEntity,
                    ChongqingStage2SettlementKinds.Proxy,
                    GetNumeric(worksheet, row, map.ProxyNetColumn),
                    !TextUtil.S(agentOrSelf).Contains("自营"));
                AddDetailIfNeeded(
                    details,
                    row,
                    customer,
                    owner,
                    intermediaryEntity,
                    ChongqingStage2SettlementKinds.Intermediary,
                    GetNumeric(worksheet, row, map.IntermediaryNetColumn),
                    true);
                AddDetailIfNeeded(
                    details,
                    row,
                    customer,
                    owner,
                    refundEntity,
                    ChongqingStage2SettlementKinds.Refund,
                    GetNumeric(worksheet, row, map.RefundNetColumn),
                    true);
            }

            return details;
        }

        private static void AddDetailIfNeeded(
            IList<ChongqingSettlementDetail> details,
            int row,
            string customer,
            string owner,
            string entity,
            string kind,
            double expectedNet,
            bool canSettle)
        {
            if (!canSettle || string.IsNullOrWhiteSpace(entity) || Math.Abs(expectedNet) <= Stage2SettlementCalculator.AmountTolerance)
            {
                return;
            }

            details.Add(new ChongqingSettlementDetail
            {
                LedgerRow = row,
                Customer = customer,
                Owner = owner,
                Entity = entity,
                Kind = kind,
                ExpectedNet = expectedNet
            });
        }

        private static void AddSummaryPaymentIssues(
            ChongqingStage2Options options,
            IList<GroupSettlementTotal> groups,
            IList<ChongqingStage2CheckIssue> issues)
        {
            if (groups.Count == 0)
            {
                return;
            }

            using (var workbook = new XLWorkbook(options.SummaryTemplatePath))
            {
                var mainSheet = FindSummaryWorksheet(workbook);
                var knownKeys = new HashSet<string>(ReadSummaryMeta(mainSheet).Select(item => SummaryKey(item.Entity, item.Kind)));
                foreach (var group in groups)
                {
                    if (knownKeys.Contains(SummaryKey(group.Entity, group.Kind))
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

        private static bool HasPaymentDecision(ChongqingStage2Options options, string entity, string kind)
        {
            var key = SummaryKey(entity, kind);
            return options.SummarySubjectDecisions
                .Where(decision => decision != null)
                .Any(decision => SummaryKey(decision.Entity, decision.SettlementKind) == key);
        }

        private static IXLWorksheet FindLedgerWorksheet(XLWorkbook workbook)
        {
            foreach (var worksheet in workbook.Worksheets)
            {
                if (FindHeaderColumn(worksheet, "电力用户名称") > 0)
                {
                    return worksheet;
                }
            }

            throw new InvalidOperationException("重庆台账中未找到表头“电力用户名称”。");
        }

        private static LedgerMap FindLedgerMap(IXLWorksheet worksheet, int month)
        {
            var start = FindMonthStartColumn(worksheet, month);
            if (CellText(worksheet.Cell(2, start)) != "总实际电量（兆瓦时）"
                || CellText(worksheet.Cell(2, start + 1)) != "实际电量（兆瓦时）"
                || CellText(worksheet.Cell(3, start + 1)) != "尖"
                || CellText(worksheet.Cell(3, start + 2)) != "峰"
                || CellText(worksheet.Cell(3, start + 3)) != "平"
                || CellText(worksheet.Cell(3, start + 4)) != "谷")
            {
                throw new InvalidOperationException("重庆台账" + month + "月月度区块表头不符合预期。");
            }

            return new LedgerMap
            {
                CustomerNameColumn = RequireHeaderColumn(worksheet, "电力用户名称"),
                ProjectDeveloperColumn = RequireHeaderColumn(worksheet, "项目开发人"),
                AgentOrSelfColumn = RequireHeaderColumn(worksheet, "代理或自营"),
                OwnerColumn = RequireHeaderColumn(worksheet, "负责人"),
                IntermediaryColumn = FindHeaderColumn(worksheet, "居间人"),
                PayeeColumn = FindHeaderColumn(worksheet, "收款人"),
                IntermediaryNetColumn = start + 12,
                RefundNetColumn = start + 21,
                ProxyNetColumn = start + 27
            };
        }

        private static int FindMonthStartColumn(IXLWorksheet worksheet, int month)
        {
            var label = month.ToString(CultureInfo.InvariantCulture) + "月";
            var lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? 1;
            for (var column = 1; column <= lastColumn; column++)
            {
                if (CellText(worksheet.Cell(1, column)) == label)
                {
                    return column;
                }
            }

            throw new InvalidOperationException("重庆台账中未找到" + label + "月度区块。");
        }

        private static int RequireHeaderColumn(IXLWorksheet worksheet, string headerText)
        {
            var column = FindHeaderColumn(worksheet, headerText);
            if (column <= 0)
            {
                throw new InvalidOperationException("重庆台账中未找到表头“" + headerText + "”。");
            }

            return column;
        }

        private static int FindHeaderColumn(IXLWorksheet worksheet, string headerText)
        {
            var lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? 1;
            for (var row = 1; row <= Math.Min(10, worksheet.LastRowUsed()?.RowNumber() ?? 1); row++)
            {
                for (var column = 1; column <= lastColumn; column++)
                {
                    if (CellText(worksheet.Cell(row, column)) == headerText)
                    {
                        return column;
                    }
                }
            }

            return 0;
        }

        private static IXLWorksheet FindSummaryWorksheet(XLWorkbook workbook)
        {
            var exact = workbook.Worksheets.FirstOrDefault(sheet => CellText(sheet.Cell("A1")).Contains("汇总")
                || sheet.Name == "汇总表");
            if (exact != null)
            {
                return exact;
            }

            return workbook.Worksheets.First();
        }

        private static List<SummaryMetaRow> ReadSummaryMeta(IXLWorksheet worksheet)
        {
            var rows = new List<SummaryMetaRow>();
            var totalRow = FindSummaryTotalRow(worksheet);
            for (var row = DataStartRow; row < totalRow; row++)
            {
                var entity = CellText(worksheet.Cell(row, 2));
                var kind = CellText(worksheet.Cell(row, 3));
                if (string.IsNullOrWhiteSpace(entity) || string.IsNullOrWhiteSpace(kind))
                {
                    continue;
                }

                rows.Add(new SummaryMetaRow { Entity = entity, Kind = kind });
            }

            return rows;
        }

        private static int FindSummaryTotalRow(IXLWorksheet worksheet)
        {
            var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? DataStartRow;
            for (var row = DataStartRow; row <= lastRow; row++)
            {
                if (CellText(worksheet.Cell(row, 1)) == "合计")
                {
                    return row;
                }
            }

            return lastRow + 1;
        }

        private static double GetNumeric(IXLWorksheet worksheet, int row, int column)
        {
            return ClosedXmlUtil.CellNumber(worksheet.Cell(row, column));
        }

        private static string CellText(IXLCell cell)
        {
            return TextUtil.S(cell.GetFormattedString());
        }

        private static string SummaryKey(string entity, string kind)
        {
            return TextUtil.CustomerKey(entity) + "|" + TextUtil.S(kind);
        }

        private sealed class LedgerMap
        {
            public int CustomerNameColumn { get; set; }
            public int ProjectDeveloperColumn { get; set; }
            public int AgentOrSelfColumn { get; set; }
            public int OwnerColumn { get; set; }
            public int IntermediaryColumn { get; set; }
            public int PayeeColumn { get; set; }
            public int IntermediaryNetColumn { get; set; }
            public int RefundNetColumn { get; set; }
            public int ProxyNetColumn { get; set; }
        }

        private sealed class ChongqingSettlementDetail
        {
            public int LedgerRow { get; set; }
            public string Customer { get; set; }
            public string Owner { get; set; }
            public string Entity { get; set; }
            public string Kind { get; set; }
            public double ExpectedNet { get; set; }
        }

        private sealed class SummaryMetaRow
        {
            public string Entity { get; set; }
            public string Kind { get; set; }
        }
    }
}
