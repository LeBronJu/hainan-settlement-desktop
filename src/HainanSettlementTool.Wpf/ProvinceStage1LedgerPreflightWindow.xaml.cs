using System.Collections.Generic;
using System.Linq;
using System.Windows;
using HainanSettlementTool.Core.Models;
using HainanSettlementTool.Core.Services;

namespace HainanSettlementTool.Wpf
{
    public partial class ProvinceStage1LedgerPreflightWindow : Window
    {
        private readonly List<ManualMatchRowViewModel> _manualMatchRows;

        public ProvinceStage1LedgerPreflightWindow(
            ProvinceStage1LedgerUpdateOptions options,
            ProvinceStage1LedgerUpdatePlan plan)
        {
            InitializeComponent();

            SummaryText.Text = ProvinceStage1Service.ProvinceName(plan.Province)
                + "；结算月份：2026年" + options.Month + "月；匹配客户："
                + plan.MatchedRows + " / " + plan.PowerCustomerRows
                + "；预检项目：" + plan.Issues.Count + " 条。";

            _manualMatchRows = BuildManualMatchRows(plan);
            ManualMatchList.ItemsSource = _manualMatchRows;
            ManualMatchCountText.Text = _manualMatchRows.Count + " 个待判断客户";
            ManualMatchPanel.Visibility = _manualMatchRows.Count > 0 ? Visibility.Visible : Visibility.Collapsed;

            IssueGroupsList.ItemsSource = BuildIssueGroups(plan);
            IssueCountText.Text = plan.Issues.Count + " 条";
        }

        public List<ProvinceStage1CustomerMatch> ManualCustomerMatches { get; private set; } =
            new List<ProvinceStage1CustomerMatch>();

        private static List<ManualMatchRowViewModel> BuildManualMatchRows(ProvinceStage1LedgerUpdatePlan plan)
        {
            var targetOptions = new List<CustomerTargetOption>
            {
                new CustomerTargetOption
                {
                    CustomerName = string.Empty,
                    DisplayText = "不匹配，本月不写入"
                }
            };
            targetOptions.AddRange(plan.LedgerOnlyCustomers
                .OrderBy(name => name)
                .Select(name => new CustomerTargetOption
                {
                    CustomerName = name,
                    DisplayText = name
                }));

            if (targetOptions.Count <= 1)
            {
                return new List<ManualMatchRowViewModel>();
            }

            return plan.PowerOnlyCustomers
                .OrderBy(name => name)
                .Select(name => new ManualMatchRowViewModel
                {
                    SourceCustomerName = name,
                    TargetOptions = targetOptions,
                    SelectedTargetName = string.Empty
                })
                .ToList();
        }

        private static List<IssueGroupViewModel> BuildIssueGroups(ProvinceStage1LedgerUpdatePlan plan)
        {
            return plan.Issues
                .GroupBy(issue => issue.Category)
                .OrderBy(group => group.Key)
                .Select(group => new IssueGroupViewModel
                {
                    Category = group.Key,
                    CountText = group.Count() + " 条",
                    Issues = group
                        .Select(issue => BuildIssueText(issue))
                        .ToList()
                })
                .ToList();
        }

        private static string BuildIssueText(ProvinceStage1LedgerUpdateIssue issue)
        {
            var customer = string.IsNullOrWhiteSpace(issue.CustomerName) ? string.Empty : "：" + issue.CustomerName;
            if (issue.Category == "电量客户不在台账")
            {
                return "[" + issue.Severity + "] " + issue.Category + customer + "；可在上方选择一个台账客户作为本次写入目标；未选择则本月不写入该客户。";
            }

            if (issue.Category == "台账客户不在电量表")
            {
                return "[" + issue.Severity + "] " + issue.Category + customer + "；可作为上方人工匹配目标；未被选择则本月不更新该台账行。";
            }

            return "[" + issue.Severity + "] " + issue.Category + customer + "；" + issue.Message;
        }

        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            var selected = _manualMatchRows
                .Where(row => !string.IsNullOrWhiteSpace(row.SelectedTargetName))
                .ToList();
            var duplicateTargets = selected
                .GroupBy(row => row.SelectedTargetName)
                .Where(group => group.Count() > 1)
                .Select(group => group.Key)
                .ToList();

            if (duplicateTargets.Count > 0)
            {
                ErrorText.Text = "同一个台账客户只能匹配一个清洗客户，请调整：" + string.Join("、", duplicateTargets);
                ErrorText.Visibility = Visibility.Visible;
                return;
            }

            ManualCustomerMatches = selected
                .Select(row => new ProvinceStage1CustomerMatch
                {
                    SourceCustomerName = row.SourceCustomerName,
                    TargetCustomerName = row.SelectedTargetName
                })
                .ToList();
            DialogResult = true;
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }

        public sealed class ManualMatchRowViewModel
        {
            public string SourceCustomerName { get; set; }
            public List<CustomerTargetOption> TargetOptions { get; set; }
            public string SelectedTargetName { get; set; }
        }

        public sealed class CustomerTargetOption
        {
            public string CustomerName { get; set; }
            public string DisplayText { get; set; }
        }

        public sealed class IssueGroupViewModel
        {
            public string Category { get; set; }
            public string CountText { get; set; }
            public List<string> Issues { get; set; }
        }
    }
}
