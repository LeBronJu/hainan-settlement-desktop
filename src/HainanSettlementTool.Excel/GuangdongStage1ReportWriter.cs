using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using HainanSettlementTool.Core.Models;
using HainanSettlementTool.Core.Services;
using Newtonsoft.Json;

namespace HainanSettlementTool.Excel
{
    internal sealed class GuangdongStage1ReportWriter
    {
        public void WriteCleanReport(
            ProvinceStage1CleanResult result,
            IEnumerable<GuangdongStage1SourceReader.GuangdongPowerAggregateRow> sourceRows)
        {
            var rows = sourceRows.ToList();
            var payload = new
            {
                province = ProvinceDisplayNames.GetName(result.Province),
                month = result.Month,
                unit = result.Unit,
                sourceSheetName = result.SourceSheetName,
                rawRows = result.RawRows,
                customerRows = result.CustomerRows,
                totalPower = rows.Sum(row => row.Total),
                coefficientConflictCustomers = rows.Count(row => row.CoefficientConflictRows > 0),
                outputWorkbookPath = result.OutputWorkbookPath,
                htmlReportPath = result.HtmlReportPath,
                warnings = result.Warnings
            };
            File.WriteAllText(
                result.ReportPath,
                JsonConvert.SerializeObject(payload, Formatting.Indented),
                Encoding.UTF8);
            File.WriteAllText(
                result.HtmlReportPath,
                ReadableHtmlReportRenderer.Render(BuildCleanDocument(result, rows)),
                Encoding.UTF8);
        }

        public void WriteLedgerReport(ProvinceStage1LedgerUpdateResult result)
        {
            var payload = new
            {
                province = ProvinceDisplayNames.GetName(result.Province),
                month = result.Month,
                unit = result.Unit,
                ledgerPath = result.LedgerPath,
                rawDetailPath = result.RawDetailPath,
                outputPowerWorkbookPath = result.OutputPowerWorkbookPath,
                outputLedgerPath = result.OutputLedgerPath,
                htmlReportPath = result.HtmlReportPath,
                ledgerCustomerRows = result.LedgerCustomerRows,
                powerCustomerRows = result.PowerCustomerRows,
                matchedRows = result.MatchedRows,
                updatedPowerRows = result.UpdatedPowerRows,
                createdCustomerRows = result.CreatedCustomerRows,
                multiMeterPointCustomers = result.MultiAccountRows,
                totalPower = result.TotalPower,
                warnings = result.Warnings,
                issues = result.Issues
            };
            File.WriteAllText(
                result.ReportPath,
                JsonConvert.SerializeObject(payload, Formatting.Indented),
                Encoding.UTF8);
            File.WriteAllText(
                result.HtmlReportPath,
                ReadableHtmlReportRenderer.Render(BuildLedgerDocument(result)),
                Encoding.UTF8);
        }

        private static ReadableReportDocument BuildCleanDocument(
            ProvinceStage1CleanResult result,
            IList<GuangdongStage1SourceReader.GuangdongPowerAggregateRow> rows)
        {
            var coefficientConflictCount = rows.Count(row => row.CoefficientConflictRows > 0);
            var hasWarnings = result.Warnings.Count > 0;
            var document = new ReadableReportDocument
            {
                Title = "广东" + result.Month + "月电量整理结果",
                PeriodLabel = "结算月份：2026年" + result.Month + "月",
                Status = hasWarnings ? ReadableReportStatus.Review : ReadableReportStatus.Success,
                StatusText = hasWarnings ? "电量表已生成，请按提示抽查" : "电量表已生成",
                StatusDetail = "请从下方“下一步”开始检查。",
                Footer = "本报告由结算自动化工具生成。一般只需查看本页；程序记录仅在需要追溯时使用。"
            };
            document.Metrics.Add(new ReadableReportMetric("原始明细", result.RawRows.ToString()));
            document.Metrics.Add(new ReadableReportMetric("客户数量", result.CustomerRows.ToString()));
            document.Metrics.Add(new ReadableReportMetric(
                "合计电量",
                result.TotalPower.ToString("0.#####", CultureInfo.InvariantCulture) + " " + result.Unit));
            document.Metrics.Add(new ReadableReportMetric("提醒事项", result.Warnings.Count.ToString()));
            document.Notices.Add(new ReadableReportNotice(
                "下一步",
                "1. 打开电量表，核对客户数量和合计电量。\n"
                + "2. 如有系数提示，抽查程序采用的首个完整系数。\n"
                + "3. 确认无误后，回到工具选择“生成本月台账”。",
                hasWarnings ? ReadableReportNoticeTone.Review : ReadableReportNoticeTone.Information));
            if (hasWarnings)
            {
                document.Notices.Add(new ReadableReportNotice(
                    "需要留意",
                    "发现 " + result.Warnings.Count + " 项系数或客户资料提醒，请打开下方“补充说明”查看。",
                    ReadableReportNoticeTone.Review));
            }

            var checks = new ReadableReportSection("重点检查", "项目", "结果");
            checks.Rows.Add(new ReadableReportRow(
                "电量合计",
                "全部明细已通过总量 = 峰 + 平 + 谷校验。"));
            checks.Rows.Add(new ReadableReportRow(
                "多个计量点",
                rows.Count(row => row.SourceRows > 1) + " 个客户已按客户编号合计。"));
            checks.Rows.Add(new ReadableReportRow(
                "峰平谷系数",
                coefficientConflictCount == 0
                    ? "未发现同一客户存在不同系数。"
                    : coefficientConflictCount + " 个客户采用明细中首个完整系数；该系数不影响代理费。"));
            document.Sections.Add(checks);

            var outputs = new ReadableReportSection("输出文件", "内容", "位置");
            outputs.Rows.Add(new ReadableReportRow("八列电量表", result.OutputWorkbookPath));
            outputs.Rows.Add(new ReadableReportRow("程序记录（一般不用打开）", result.ReportPath));
            document.Sections.Add(outputs);
            AddWarnings(document, result.Warnings);
            return document;
        }

        private static ReadableReportDocument BuildLedgerDocument(
            ProvinceStage1LedgerUpdateResult result)
        {
            var groups = ProvinceStage1ReviewGuide.Build(result.Issues);
            var focusGroups = groups.Where(group => group.NeedsAttention).ToList();
            var ledgerOnlyCount = System.Math.Max(0, result.LedgerCustomerRows - result.MatchedRows);
            var ledgerReviewStep = result.CreatedCustomerRows > 0
                ? "2. 打开台账，按“重点检查”补齐新增客户资料，并核对名称/编号提示。\n"
                : "2. 打开台账，按“重点检查”核对名称/编号提示。\n";
            var document = new ReadableReportDocument
            {
                Title = "广东" + result.Month + "月台账生成结果",
                PeriodLabel = "结算月份：2026年" + result.Month + "月",
                Status = focusGroups.Count > 0 ? ReadableReportStatus.Review : ReadableReportStatus.Success,
                StatusText = focusGroups.Count > 0 ? "已生成，请完成重点检查" : "已生成，可以继续下一步",
                StatusDetail = "请从下方“下一步”开始检查。",
                Footer = "本报告由结算自动化工具生成。一般只需查看本页；程序记录仅在需要追溯时使用。"
            };
            document.Metrics.Add(new ReadableReportMetric("来源客户", result.PowerCustomerRows.ToString()));
            document.Metrics.Add(new ReadableReportMetric("写入客户", result.UpdatedPowerRows.ToString()));
            document.Metrics.Add(new ReadableReportMetric("新增客户", result.CreatedCustomerRows.ToString()));
            document.Metrics.Add(new ReadableReportMetric("本月无电量", ledgerOnlyCount.ToString()));
            document.Notices.Add(new ReadableReportNotice(
                "下一步",
                "1. 打开八列电量表，核对客户数量和合计电量。\n"
                + ledgerReviewStep
                + "3. 抽查本月无电量客户的总量、峰、平、谷均为 0，原有系数仍保留。\n"
                + "4. 确认无误后，再进行人工资料整理和后续结算。",
                focusGroups.Count > 0
                    ? ReadableReportNoticeTone.Review
                    : ReadableReportNoticeTone.Information));

            var focus = new ReadableReportSection("重点检查", "事项", "数量", "客户预览", "你要做什么");
            foreach (var group in focusGroups)
            {
                focus.Rows.Add(new ReadableReportRow(
                    group.Title,
                    group.CountText,
                    group.CustomerPreview,
                    group.ActionText));
            }
            focus.EmptyMessage = "没有需要额外处理的项目。";
            document.Sections.Add(focus);

            var automatic = new ReadableReportSection("程序已经处理", "项目", "结果");
            automatic.Rows.Add(new ReadableReportRow(
                "本月无电量客户",
                ledgerOnlyCount + " 个客户的总量、峰、平、谷已写入 0，原有系数保留。"));
            automatic.Rows.Add(new ReadableReportRow(
                "多个计量点",
                result.MultiAccountRows + " 个客户已按客户编号合计。"));
            automatic.Rows.Add(new ReadableReportRow(
                "写入方式",
                "原台账未覆盖；结果写入新的台账文件。"));
            document.Sections.Add(automatic);

            var outputs = new ReadableReportSection("输出文件", "内容", "位置");
            outputs.Rows.Add(new ReadableReportRow("八列电量表", result.OutputPowerWorkbookPath));
            outputs.Rows.Add(new ReadableReportRow("本月台账", result.OutputLedgerPath));
            outputs.Rows.Add(new ReadableReportRow("程序记录（一般不用打开）", result.ReportPath));
            document.Sections.Add(outputs);

            var details = new ReadableReportSection("完整提示明细", "类型", "客户", "说明")
            {
                IsCollapsed = true,
                EmptyMessage = "没有其它提示。"
            };
            foreach (var issue in result.Issues)
            {
                details.Rows.Add(new ReadableReportRow(
                    string.IsNullOrWhiteSpace(issue.Category) ? issue.Kind : issue.Category,
                    issue.CustomerName,
                    issue.Message));
            }
            document.Sections.Add(details);
            AddWarnings(document, result.Warnings);
            return document;
        }

        private static void AddWarnings(
            ReadableReportDocument document,
            IEnumerable<string> warnings)
        {
            var rows = warnings
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Distinct()
                .ToList();
            if (rows.Count == 0)
            {
                return;
            }

            var section = new ReadableReportSection("补充说明", "内容")
            {
                IsCollapsed = true
            };
            foreach (var warning in rows)
            {
                section.Rows.Add(new ReadableReportRow(warning));
            }
            document.Sections.Add(section);
        }
    }
}
