using System;
using System.Collections.Generic;
using System.Linq;
using HainanSettlementTool.Core.Models;
using HainanSettlementTool.Core.Services;

namespace HainanSettlementTool.Wpf
{
    internal static class ProvinceStage1PreflightPresentationAdapter
    {
        public static ProvinceStage1PreflightDialogViewModel Create(
            ProvinceStage1LedgerUpdateOptions options,
            ProvinceStage1LedgerUpdatePlan plan)
        {
            if (options == null)
            {
                throw new ArgumentNullException(nameof(options));
            }

            if (plan == null)
            {
                throw new ArgumentNullException(nameof(plan));
            }

            var groups = ProvinceStage1ReviewGuide.Build(plan.Issues);
            var focusGroups = groups.Where(group => group.NeedsAttention).ToList();
            var automaticGroups = groups.Where(group => !group.NeedsAttention).ToList();
            var provinceName = ProvinceDisplayNames.GetName(plan.Province);
            return new ProvinceStage1PreflightDialogViewModel
            {
                Title = "生成" + provinceName + "本月台账前确认",
                Heading = focusGroups.Count > 0 ? "预检报告" : "检查完成，可以生成",
                SummaryText = "2026年" + options.Month + "月",
                GuidanceText = BuildGuidance(plan, focusGroups.Count),
                ConfirmButtonText = plan.Province == ProvinceCode.Guangdong
                    ? "确认并生成本月台账"
                    : "确认并生成",
                Metrics = BuildMetrics(plan, focusGroups.Count),
                FocusGroups = focusGroups,
                AutomaticItems = BuildAutomaticItems(plan, automaticGroups),
                DetailGroups = BuildDetailGroups(plan),
                DetailsHeaderText = "查看完整明细（" + plan.Issues.Count + " 项）",
                FocusCountText = focusGroups.Count + " 类需要检查"
            };
        }

        private static string BuildGuidance(
            ProvinceStage1LedgerUpdatePlan plan,
            int focusGroupCount)
        {
            if (plan.Province == ProvinceCode.Guangdong)
            {
                return focusGroupCount > 0
                    ? "请先看“重点检查”。程序会按客户编号处理，不会改原台账；生成后可在检查报告里按步骤复核。"
                    : "没有需要逐户决定的事项。程序会按客户编号处理，不会改原台账；生成后请按检查报告核对结果。";
            }

            return "需要匹配的客户请逐一选择；其余事项已经按类型归纳。生成结果不会覆盖原台账。";
        }

        private static List<ProvinceStage1PreflightMetricViewModel> BuildMetrics(
            ProvinceStage1LedgerUpdatePlan plan,
            int focusGroupCount)
        {
            return new List<ProvinceStage1PreflightMetricViewModel>
            {
                new ProvinceStage1PreflightMetricViewModel("电量表客户", plan.PowerCustomerRows.ToString()),
                new ProvinceStage1PreflightMetricViewModel("台账已有", plan.MatchedRows.ToString()),
                new ProvinceStage1PreflightMetricViewModel("新增客户", plan.CreatedCustomerRows.ToString()),
                new ProvinceStage1PreflightMetricViewModel("重点检查", focusGroupCount + " 类")
            };
        }

        private static List<string> BuildAutomaticItems(
            ProvinceStage1LedgerUpdatePlan plan,
            IEnumerable<ProvinceStage1ReviewGroup> automaticGroups)
        {
            var items = new List<string>();
            if (plan.MatchedRows > 0)
            {
                items.Add(plan.MatchedRows + " 个客户已在台账找到，将写入本月电量。");
            }

            if (plan.MultiAccountRows > 0)
            {
                items.Add(plan.MultiAccountRows + " 个客户有多个计量点，电量已经合计。");
            }

            if (plan.MissingInPowerRows > 0)
            {
                items.Add(plan.Province == ProvinceCode.Guangdong
                    ? plan.MissingInPowerRows + " 个台账客户本月无电量，四项电量将写入 0，原有系数保留。"
                    : plan.MissingInPowerRows + " 个台账客户本月无来源电量，旧电量不会沿用。");
            }

            items.AddRange(automaticGroups
                .Where(group => group.Kind != ProvinceStage1LedgerUpdateIssueKinds.LedgerCustomerMissingInPower
                    && group.Kind != ProvinceStage1LedgerUpdateIssueKinds.MultiAccountCustomer)
                .Select(group => group.Title + "：" + group.ActionText));
            if (items.Count == 0)
            {
                items.Add("来源客户和台账客户均已正常对应。");
            }

            return items;
        }

        private static List<ProvinceStage1PreflightDetailGroupViewModel> BuildDetailGroups(
            ProvinceStage1LedgerUpdatePlan plan)
        {
            return plan.Issues
                .Where(issue => issue != null)
                .GroupBy(issue => string.IsNullOrWhiteSpace(issue.Category) ? issue.Kind : issue.Category)
                .OrderBy(group => group.Key, StringComparer.CurrentCulture)
                .Select(group => new ProvinceStage1PreflightDetailGroupViewModel
                {
                    Category = group.Key,
                    CountText = group.Count() + " 项",
                    Issues = group.Select(BuildDetailText).ToList()
                })
                .ToList();
        }

        private static string BuildDetailText(ProvinceStage1LedgerUpdateIssue issue)
        {
            var prefix = string.IsNullOrWhiteSpace(issue.CustomerName)
                ? string.Empty
                : issue.CustomerName + "：";
            return prefix + issue.Message;
        }
    }

    internal sealed class ProvinceStage1PreflightDialogViewModel
    {
        public string Title { get; set; }
        public string Heading { get; set; }
        public string SummaryText { get; set; }
        public string GuidanceText { get; set; }
        public string ConfirmButtonText { get; set; }
        public string FocusCountText { get; set; }
        public string DetailsHeaderText { get; set; }
        public List<ProvinceStage1PreflightMetricViewModel> Metrics { get; set; }
        public List<ProvinceStage1ReviewGroup> FocusGroups { get; set; }
        public List<string> AutomaticItems { get; set; }
        public List<ProvinceStage1PreflightDetailGroupViewModel> DetailGroups { get; set; }
    }

    internal sealed class ProvinceStage1PreflightMetricViewModel
    {
        public ProvinceStage1PreflightMetricViewModel(string label, string value)
        {
            Label = label;
            Value = value;
        }

        public string Label { get; }
        public string Value { get; }
    }

    internal sealed class ProvinceStage1PreflightDetailGroupViewModel
    {
        public string Category { get; set; }
        public string CountText { get; set; }
        public List<string> Issues { get; set; }
    }
}
