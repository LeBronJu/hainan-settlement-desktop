using System;
using System.Collections.Generic;
using System.Linq;
using HainanSettlementTool.Core.Models;

namespace HainanSettlementTool.Core.Services
{
    public static class ProvinceStage1ReviewGuide
    {
        private const int CustomerPreviewLimit = 5;

        public static List<ProvinceStage1ReviewGroup> Build(
            IEnumerable<ProvinceStage1LedgerUpdateIssue> issues)
        {
            return (issues ?? Enumerable.Empty<ProvinceStage1LedgerUpdateIssue>())
                .Where(issue => issue != null)
                .GroupBy(
                    issue => string.IsNullOrWhiteSpace(issue.Kind) ? issue.Category : issue.Kind,
                    StringComparer.Ordinal)
                .Select(BuildGroup)
                .OrderBy(group => group.NeedsAttention ? 0 : 1)
                .ThenBy(group => group.DisplayOrder)
                .ThenBy(group => group.Title, StringComparer.CurrentCulture)
                .ToList();
        }

        private static ProvinceStage1ReviewGroup BuildGroup(
            IGrouping<string, ProvinceStage1LedgerUpdateIssue> source)
        {
            var rows = source.ToList();
            var first = rows[0];
            var kind = string.IsNullOrWhiteSpace(first.Kind) ? first.Category : first.Kind;
            var customers = rows
                .Select(issue => issue.CustomerName)
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Distinct(StringComparer.CurrentCulture)
                .OrderBy(value => value, StringComparer.CurrentCulture)
                .ToList();
            return new ProvinceStage1ReviewGroup
            {
                Kind = kind,
                Title = TitleFor(kind, first.Category),
                ActionText = ActionFor(kind),
                NeedsAttention = NeedsAttention(kind, rows),
                DisplayOrder = DisplayOrderFor(kind),
                IssueCount = rows.Count,
                CustomerCount = customers.Count,
                CustomerPreview = BuildCustomerPreview(customers),
                CustomerNames = customers
            };
        }

        private static bool NeedsAttention(
            string kind,
            IEnumerable<ProvinceStage1LedgerUpdateIssue> issues)
        {
            switch (kind)
            {
                case ProvinceStage1LedgerUpdateIssueKinds.LedgerCustomerMissingInPower:
                case ProvinceStage1LedgerUpdateIssueKinds.MultiAccountCustomer:
                case ProvinceStage1LedgerUpdateIssueKinds.ManualMatchedCustomer:
                    return false;
                case ProvinceStage1LedgerUpdateIssueKinds.CreatedCustomer:
                case ProvinceStage1LedgerUpdateIssueKinds.ExistingPowerDifference:
                case ProvinceStage1LedgerUpdateIssueKinds.PowerCustomerMissingInLedger:
                case ProvinceStage1LedgerUpdateIssueKinds.SkippedPowerCustomer:
                case ProvinceStage1LedgerUpdateIssueKinds.PossibleAlias:
                case ProvinceStage1LedgerUpdateIssueKinds.CustomerNameMismatch:
                case ProvinceStage1LedgerUpdateIssueKinds.CoefficientConflict:
                case ProvinceStage1LedgerUpdateIssueKinds.MonthMismatch:
                    return true;
                default:
                    return issues.Any(issue =>
                        !string.Equals(issue.Severity, "提示", StringComparison.Ordinal));
            }
        }

        private static string TitleFor(string kind, string fallback)
        {
            switch (kind)
            {
                case ProvinceStage1LedgerUpdateIssueKinds.MonthMismatch:
                    return "月份需要确认";
                case ProvinceStage1LedgerUpdateIssueKinds.CreatedCustomer:
                case ProvinceStage1LedgerUpdateIssueKinds.PowerCustomerMissingInLedger:
                    return "新增客户";
                case ProvinceStage1LedgerUpdateIssueKinds.ExistingPowerDifference:
                    return "本次电量将覆盖旧值";
                case ProvinceStage1LedgerUpdateIssueKinds.PossibleAlias:
                    return "同名但客户编号不同";
                case ProvinceStage1LedgerUpdateIssueKinds.CustomerNameMismatch:
                    return "客户编号相同但名称不同";
                case ProvinceStage1LedgerUpdateIssueKinds.CoefficientConflict:
                    return "多计量点峰平谷系数不同客户";
                case ProvinceStage1LedgerUpdateIssueKinds.LedgerCustomerMissingInPower:
                    return "本月无电量客户";
                case ProvinceStage1LedgerUpdateIssueKinds.MultiAccountCustomer:
                    return "多个计量点已合计";
                case ProvinceStage1LedgerUpdateIssueKinds.ManualMatchedCustomer:
                    return "已按你的选择匹配";
                case ProvinceStage1LedgerUpdateIssueKinds.SkippedPowerCustomer:
                    return "本月不写入的客户";
                default:
                    return string.IsNullOrWhiteSpace(fallback) ? "其他检查事项" : fallback;
            }
        }

        private static string ActionFor(string kind)
        {
            switch (kind)
            {
                case ProvinceStage1LedgerUpdateIssueKinds.MonthMismatch:
                    return "请确认界面月份和文件月份一致；不一致时返回修改。";
                case ProvinceStage1LedgerUpdateIssueKinds.CreatedCustomer:
                case ProvinceStage1LedgerUpdateIssueKinds.PowerCustomerMissingInLedger:
                    return "生成后请补齐合同、税务、代理关系、负责人等人工资料。";
                case ProvinceStage1LedgerUpdateIssueKinds.ExistingPowerDifference:
                    return "请确认本次零售结算明细是应采用的最新电量。";
                case ProvinceStage1LedgerUpdateIssueKinds.PossibleAlias:
                    return "请核对是否确为不同客户；程序不会自动合并。";
                case ProvinceStage1LedgerUpdateIssueKinds.CustomerNameMismatch:
                    return "请核对名称差异；程序按客户编号写入，不会改台账名称。";
                case ProvinceStage1LedgerUpdateIssueKinds.CoefficientConflict:
                    return "不影响代理费结算，但建议检查台账。";
                case ProvinceStage1LedgerUpdateIssueKinds.LedgerCustomerMissingInPower:
                    return "总量、峰、平、谷写入 0，原有峰平谷系数保留，无需逐户操作。";
                case ProvinceStage1LedgerUpdateIssueKinds.MultiAccountCustomer:
                    return "同一客户的多个计量点已经相加，无需重复处理。";
                case ProvinceStage1LedgerUpdateIssueKinds.ManualMatchedCustomer:
                    return "本次按已确认的客户对应关系写入。";
                case ProvinceStage1LedgerUpdateIssueKinds.SkippedPowerCustomer:
                    return "这些客户本月不会写入台账，请确认这是有意选择。";
                default:
                    return "请按完整明细中的说明核对。";
            }
        }

        private static int DisplayOrderFor(string kind)
        {
            switch (kind)
            {
                case ProvinceStage1LedgerUpdateIssueKinds.MonthMismatch:
                    return 0;
                case ProvinceStage1LedgerUpdateIssueKinds.CreatedCustomer:
                case ProvinceStage1LedgerUpdateIssueKinds.PowerCustomerMissingInLedger:
                    return 10;
                case ProvinceStage1LedgerUpdateIssueKinds.PossibleAlias:
                case ProvinceStage1LedgerUpdateIssueKinds.CustomerNameMismatch:
                    return 20;
                case ProvinceStage1LedgerUpdateIssueKinds.ExistingPowerDifference:
                    return 30;
                case ProvinceStage1LedgerUpdateIssueKinds.CoefficientConflict:
                    return 40;
                case ProvinceStage1LedgerUpdateIssueKinds.LedgerCustomerMissingInPower:
                    return 50;
                default:
                    return 100;
            }
        }

        private static string BuildCustomerPreview(IList<string> customers)
        {
            if (customers.Count == 0)
            {
                return "全批次";
            }

            var preview = string.Join("、", customers.Take(CustomerPreviewLimit));
            return customers.Count > CustomerPreviewLimit
                ? preview + " 等 " + customers.Count + " 个"
                : preview;
        }
    }

    public sealed class ProvinceStage1ReviewGroup
    {
        public string Kind { get; set; }
        public string Title { get; set; }
        public string ActionText { get; set; }
        public bool NeedsAttention { get; set; }
        public int DisplayOrder { get; set; }
        public int IssueCount { get; set; }
        public int CustomerCount { get; set; }
        public string CustomerPreview { get; set; }
        public List<string> CustomerNames { get; set; } = new List<string>();

        public string CountText
        {
            get { return IssueCount + " 项"; }
        }
    }
}
