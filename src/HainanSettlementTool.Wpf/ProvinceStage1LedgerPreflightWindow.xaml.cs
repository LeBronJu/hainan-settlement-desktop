using System.Collections.Generic;
using System.Linq;
using System.Windows;
using HainanSettlementTool.Core.Models;
using HainanSettlementTool.Core.Services;

namespace HainanSettlementTool.Wpf
{
    public partial class ProvinceStage1LedgerPreflightWindow : Window
    {
        private const string SkipManualMatchValue = "__NO_WRITE__";
        private readonly List<ManualMatchRowViewModel> _manualMatchRows;

        public ProvinceStage1LedgerPreflightWindow(
            ProvinceStage1LedgerUpdateOptions options,
            ProvinceStage1LedgerUpdatePlan plan)
        {
            InitializeComponent();

            var manualMatchingIssueCount = CountManualMatchingIssues(plan);
            var issueGroups = BuildIssueGroups(plan);
            var otherIssueCount = issueGroups.Sum(group => group.Issues.Count);

            SummaryText.Text = ProvinceDisplayNames.GetName(plan.Province)
                + "；结算月份：2026年" + options.Month + "月；匹配客户："
                + plan.MatchedRows + " / " + plan.PowerCustomerRows
                + "；客户匹配项目：" + manualMatchingIssueCount + " 条；其它预检项目：" + otherIssueCount + " 条。";

            _manualMatchRows = BuildManualMatchRows(plan);
            PowerOnlySummaryText.Text = BuildCustomerSummary(plan.PowerOnlyCustomers);
            LedgerOnlySummaryText.Text = BuildCustomerSummary(plan.LedgerOnlyCustomers);
            ManualMatchList.ItemsSource = _manualMatchRows;
            ManualMatchCountText.Text = _manualMatchRows.Count + " 个待手动选择";
            ManualMatchPanel.Visibility = manualMatchingIssueCount > 0 ? Visibility.Visible : Visibility.Collapsed;
            ManualMatchList.Visibility = _manualMatchRows.Count > 0 ? Visibility.Visible : Visibility.Collapsed;
            NoManualMatchRowsText.Visibility = _manualMatchRows.Count > 0 ? Visibility.Collapsed : Visibility.Visible;

            IssueGroupsList.ItemsSource = issueGroups;
            IssueCountText.Text = otherIssueCount + " 条";
            OtherIssuePanel.Visibility = otherIssueCount > 0 ? Visibility.Visible : Visibility.Collapsed;
        }

        public List<ProvinceStage1CustomerMatch> ManualCustomerMatches { get; private set; } =
            new List<ProvinceStage1CustomerMatch>();

        private static List<ManualMatchRowViewModel> BuildManualMatchRows(ProvinceStage1LedgerUpdatePlan plan)
        {
            var targetOptions = new List<CustomerTargetOption>
            {
                new CustomerTargetOption
                {
                    CustomerName = SkipManualMatchValue,
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
                    SelectedTargetName = null
                })
                .ToList();
        }

        private static List<IssueGroupViewModel> BuildIssueGroups(ProvinceStage1LedgerUpdatePlan plan)
        {
            return plan.Issues
                .Where(issue => !IsManualMatchingIssue(issue))
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

        private static int CountManualMatchingIssues(ProvinceStage1LedgerUpdatePlan plan)
        {
            return plan.Issues.Count(IsManualMatchingIssue);
        }

        private static bool IsManualMatchingIssue(ProvinceStage1LedgerUpdateIssue issue)
        {
            return issue.Kind == ProvinceStage1LedgerUpdateIssueKinds.PowerCustomerMissingInLedger
                || issue.Kind == ProvinceStage1LedgerUpdateIssueKinds.LedgerCustomerMissingInPower
                || issue.Category == "电量客户不在台账"
                || issue.Category == "台账客户不在电量表";
        }

        private static string BuildCustomerSummary(IList<string> customers)
        {
            if (customers == null || customers.Count == 0)
            {
                return "无";
            }

            return customers.Count + " 个：" + string.Join("、", customers.OrderBy(name => name));
        }

        private static string BuildIssueText(ProvinceStage1LedgerUpdateIssue issue)
        {
            var customer = string.IsNullOrWhiteSpace(issue.CustomerName) ? string.Empty : "：" + issue.CustomerName;
            return "[" + issue.Severity + "] " + issue.Category + customer + "；" + issue.Message;
        }

        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            var unselectedRows = _manualMatchRows
                .Where(row => string.IsNullOrWhiteSpace(row.SelectedTargetName))
                .ToList();
            if (unselectedRows.Count > 0)
            {
                ErrorText.Text = "请先为以下客户选择台账客户名称，或明确选择“不匹配，本月不写入”：" + BuildShortCustomerList(unselectedRows.Select(row => row.SourceCustomerName));
                ErrorText.Visibility = Visibility.Visible;
                return;
            }

            var selected = _manualMatchRows
                .Where(row => row.SelectedTargetName != SkipManualMatchValue)
                .ToList();
            var duplicateTargets = selected
                .GroupBy(row => row.SelectedTargetName)
                .Where(group => group.Count() > 1)
                .Select(group => group.Key)
                .ToList();

            if (duplicateTargets.Count > 0)
            {
                ErrorText.Text = "同一个台账客户只能匹配一个待匹配客户，请调整：" + string.Join("、", duplicateTargets);
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

        private static string BuildShortCustomerList(IEnumerable<string> customers)
        {
            var names = customers.Take(5).ToList();
            var suffix = customers.Skip(5).Any() ? " 等" : string.Empty;
            return string.Join("、", names) + suffix;
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

            public override string ToString()
            {
                return DisplayText;
            }
        }

        public sealed class IssueGroupViewModel
        {
            public string Category { get; set; }
            public string CountText { get; set; }
            public List<string> Issues { get; set; }
        }
    }
}
