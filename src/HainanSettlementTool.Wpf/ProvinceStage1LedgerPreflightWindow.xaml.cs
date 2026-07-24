using System.Collections.Generic;
using System.Linq;
using System.Windows;
using HainanSettlementTool.Core.Models;
using HainanSettlementTool.Core.Services;

namespace HainanSettlementTool.Wpf
{
    public partial class ProvinceStage1LedgerPreflightWindow : Window
    {
        private const string CreateNewCustomerValue = "__CREATE_NEW__";
        private const string SkipManualMatchValue = "__NO_WRITE__";
        private readonly List<ManualMatchRowViewModel> _manualMatchRows;

        public ProvinceStage1LedgerPreflightWindow(
            ProvinceStage1LedgerUpdateOptions options,
            ProvinceStage1LedgerUpdatePlan plan)
        {
            InitializeComponent();

            var presentation = ProvinceStage1PreflightPresentationAdapter.Create(options, plan);
            var manualMatchingIssueCount = CountManualMatchingIssues(plan);
            Title = presentation.Title;
            HeadingText.Text = presentation.Heading;
            SummaryText.Text = presentation.SummaryText;
            GuidanceText.Text = presentation.GuidanceText;
            ConfirmButton.Content = presentation.ConfirmButtonText;
            SummaryMetricsList.ItemsSource = presentation.Metrics;
            FocusGroupsList.ItemsSource = presentation.FocusGroups;
            FocusCountText.Text = presentation.FocusCountText;
            FocusPanel.Visibility = presentation.FocusGroups.Count > 0
                ? Visibility.Visible
                : Visibility.Collapsed;
            AutomaticItemsList.ItemsSource = presentation.AutomaticItems;
            AutomaticPanel.Visibility = presentation.AutomaticItems.Count > 0
                ? Visibility.Visible
                : Visibility.Collapsed;
            IssueGroupsList.ItemsSource = presentation.DetailGroups;
            TechnicalDetailsHeaderText.Text = presentation.DetailsHeaderText;
            TechnicalDetailsExpander.Visibility = presentation.DetailGroups.Count > 0
                ? Visibility.Visible
                : Visibility.Collapsed;

            _manualMatchRows = BuildManualMatchRows(plan);
            PowerOnlySummaryText.Text = BuildCustomerSummary(plan.PowerOnlyCustomers);
            LedgerOnlySummaryText.Text = BuildCustomerSummary(plan.LedgerOnlyCustomers);
            ManualMatchList.ItemsSource = _manualMatchRows;
            ManualMatchCountText.Text = _manualMatchRows.Count + " 个待手动选择";
            ManualMatchPanel.Visibility = manualMatchingIssueCount > 0 ? Visibility.Visible : Visibility.Collapsed;
            ManualMatchList.Visibility = _manualMatchRows.Count > 0 ? Visibility.Visible : Visibility.Collapsed;
            NoManualMatchRowsText.Visibility = _manualMatchRows.Count > 0 ? Visibility.Collapsed : Visibility.Visible;
        }

        public List<ProvinceStage1CustomerMatch> ManualCustomerMatches { get; private set; } =
            new List<ProvinceStage1CustomerMatch>();

        public List<ProvinceStage1CustomerDecision> CustomerDecisions { get; private set; } =
            new List<ProvinceStage1CustomerDecision>();

        private static List<ManualMatchRowViewModel> BuildManualMatchRows(ProvinceStage1LedgerUpdatePlan plan)
        {
            var targetOptions = new List<CustomerTargetOption>
            {
                new CustomerTargetOption
                {
                    Value = CreateNewCustomerValue,
                    CustomerName = null,
                    DecisionKind = ProvinceStage1CustomerDecisionKind.CreateNew,
                    DisplayText = "新增客户到台账"
                },
                new CustomerTargetOption
                {
                    Value = SkipManualMatchValue,
                    CustomerName = SkipManualMatchValue,
                    DecisionKind = ProvinceStage1CustomerDecisionKind.SkipWrite,
                    DisplayText = "不匹配，本月不写入"
                }
            };
            targetOptions.AddRange(plan.LedgerOnlyCustomers
                .OrderBy(name => name)
                .Select(name => new CustomerTargetOption
                {
                    Value = name,
                    CustomerName = name,
                    DecisionKind = ProvinceStage1CustomerDecisionKind.MatchExisting,
                    DisplayText = name
                }));

            return plan.PowerOnlyCustomers
                .OrderBy(name => name)
                .Select(name => new ManualMatchRowViewModel
                {
                    SourceCustomerName = name,
                    TargetOptions = targetOptions,
                    SelectedTargetValue = null
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
                || issue.Category == "电量客户不在台账";
        }

        private static string BuildCustomerSummary(IList<string> customers)
        {
            if (customers == null || customers.Count == 0)
            {
                return "无";
            }

            var ordered = customers.OrderBy(name => name).ToList();
            var preview = string.Join("、", ordered.Take(5));
            return customers.Count + " 个：" + preview + (ordered.Count > 5 ? " 等" : string.Empty);
        }

        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            var unselectedRows = _manualMatchRows
                .Where(row => string.IsNullOrWhiteSpace(row.SelectedTargetValue))
                .ToList();
            if (unselectedRows.Count > 0)
            {
                ErrorText.Text = "请先为以下客户选择台账客户名称，或明确选择“不匹配，本月不写入”：" + BuildShortCustomerList(unselectedRows.Select(row => row.SourceCustomerName));
                ErrorText.Visibility = Visibility.Visible;
                return;
            }

            var selectedOptions = _manualMatchRows
                .Select(row => new
                {
                    Row = row,
                    Option = row.TargetOptions.FirstOrDefault(option => option.Value == row.SelectedTargetValue)
                })
                .ToList();

            if (selectedOptions.Any(item => item.Option == null))
            {
                ErrorText.Text = "存在无法识别的客户处理决定，请重新选择。";
                ErrorText.Visibility = Visibility.Visible;
                return;
            }

            var duplicateTargets = selectedOptions
                .Where(item => item.Option.DecisionKind == ProvinceStage1CustomerDecisionKind.MatchExisting)
                .GroupBy(item => item.Option.CustomerName)
                .Where(group => group.Count() > 1)
                .Select(group => group.Key)
                .ToList();

            if (duplicateTargets.Count > 0)
            {
                ErrorText.Text = "同一个台账客户只能匹配一个待匹配客户，请调整：" + string.Join("、", duplicateTargets);
                ErrorText.Visibility = Visibility.Visible;
                return;
            }

            CustomerDecisions = selectedOptions
                .Select(item => new ProvinceStage1CustomerDecision
                {
                    SourceCustomerName = item.Row.SourceCustomerName,
                    DecisionKind = item.Option.DecisionKind,
                    TargetCustomerName = item.Option.DecisionKind == ProvinceStage1CustomerDecisionKind.MatchExisting
                        ? item.Option.CustomerName
                        : null
                })
                .ToList();
            ManualCustomerMatches = CustomerDecisions
                .Where(decision => decision.DecisionKind == ProvinceStage1CustomerDecisionKind.MatchExisting)
                .Select(decision => new ProvinceStage1CustomerMatch
                {
                    SourceCustomerName = decision.SourceCustomerName,
                    TargetCustomerName = decision.TargetCustomerName
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
            public string SelectedTargetValue { get; set; }
        }

        public sealed class CustomerTargetOption
        {
            public string Value { get; set; }
            public string CustomerName { get; set; }
            public ProvinceStage1CustomerDecisionKind DecisionKind { get; set; }
            public string DisplayText { get; set; }

            public override string ToString()
            {
                return DisplayText;
            }
        }
    }
}
