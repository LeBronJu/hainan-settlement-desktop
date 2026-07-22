using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;

namespace HainanSettlementTool.Wpf
{
    public partial class Stage2PreflightWindow : Window
    {
        private readonly Stage2PreflightDialogViewModel _viewModel;

        internal Stage2PreflightWindow(Stage2PreflightDialogViewModel viewModel)
        {
            _viewModel = viewModel ?? throw new ArgumentNullException(nameof(viewModel));
            InitializeComponent();
            DataContext = viewModel;
            UpdateConfirmButtonState();
        }

        internal List<Stage2PreflightPaymentDecision> PaymentDecisions { get; private set; } =
            new List<Stage2PreflightPaymentDecision>();

        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            if (!_viewModel.CanConfirm)
            {
                ErrorText.Text = _viewModel.BlockingMessage;
                ErrorText.Visibility = Visibility.Visible;
                return;
            }

            var missing = _viewModel.IssueRows
                .Where(row => row.RequiresPaymentPartySelection
                    && string.IsNullOrWhiteSpace(row.SelectedPaymentParty))
                .Select(BuildSubjectText)
                .ToList();
            if (missing.Count > 0)
            {
                ErrorText.Text = "请先为以下汇总主体选择支付方：" + BuildShortSubjectList(missing);
                ErrorText.Visibility = Visibility.Visible;
                return;
            }

            PaymentDecisions = _viewModel.IssueRows
                .Where(row => row.RequiresPaymentPartySelection)
                .Select(row => new Stage2PreflightPaymentDecision
                {
                    SettlementKind = row.SettlementKind,
                    Entity = row.Entity,
                    PaymentParty = row.SelectedPaymentParty
                })
                .ToList();
            DialogResult = true;
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }

        private void PaymentParty_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ErrorText.Visibility = Visibility.Collapsed;
            Dispatcher.BeginInvoke(
                new Action(UpdateConfirmButtonState),
                DispatcherPriority.DataBind);
        }

        private void UpdateConfirmButtonState()
        {
            if (ConfirmButton == null)
            {
                return;
            }

            ConfirmButton.IsEnabled = _viewModel.CanConfirm
                && _viewModel.IssueRows
                    .Where(row => row.RequiresPaymentPartySelection)
                    .All(row => !string.IsNullOrWhiteSpace(row.SelectedPaymentParty));
        }

        private static string BuildSubjectText(Stage2PreflightIssueItemViewModel row)
        {
            var parts = new[] { row.SettlementKind, row.Entity }
                .Where(part => !string.IsNullOrWhiteSpace(part));
            var text = string.Join(" ", parts);
            return string.IsNullOrWhiteSpace(text) ? "未命名主体" : text;
        }

        private static string BuildShortSubjectList(IEnumerable<string> subjects)
        {
            var names = subjects.Take(5).ToList();
            var suffix = subjects.Skip(5).Any() ? " 等" : string.Empty;
            return string.Join("、", names) + suffix;
        }
    }
}
