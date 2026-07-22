using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using HainanSettlementTool.Core.Services;

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

        internal List<Stage2PreflightTemplateDecision> TemplateDecisions { get; private set; } =
            new List<Stage2PreflightTemplateDecision>();

        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            if (!_viewModel.CanConfirm)
            {
                ErrorText.Text = _viewModel.BlockingMessage;
                ErrorText.Visibility = Visibility.Visible;
                return;
            }

            var missingPayments = _viewModel.IssueRows
                .Where(row => row.RequiresPaymentPartySelection
                    && string.IsNullOrWhiteSpace(row.SelectedPaymentParty))
                .Select(BuildSubjectText)
                .Distinct(StringComparer.CurrentCulture)
                .ToList();
            var missingTemplates = _viewModel.IssueRows
                .Where(row => row.RequiresTemplateSelection
                    && string.IsNullOrWhiteSpace(row.SelectedTemplatePath))
                .Select(BuildSubjectText)
                .Distinct(StringComparer.CurrentCulture)
                .ToList();
            if (missingPayments.Count > 0 || missingTemplates.Count > 0)
            {
                ErrorText.Text = BuildMissingDecisionMessage(missingPayments, missingTemplates);
                ErrorText.Visibility = Visibility.Visible;
                return;
            }

            var invalidPayment = _viewModel.IssueRows.FirstOrDefault(row =>
                row.RequiresPaymentPartySelection && !HasValidPaymentSelection(row));
            var invalidTemplate = _viewModel.IssueRows.FirstOrDefault(row =>
                row.RequiresTemplateSelection && !HasValidTemplateSelection(row));
            if (invalidPayment != null || invalidTemplate != null)
            {
                ErrorText.Text = invalidPayment != null
                    ? "支付方选择已失效，请为" + BuildSubjectText(invalidPayment) + "重新选择。"
                    : "分表模板选择已失效，请为" + BuildSubjectText(invalidTemplate) + "重新选择。";
                ErrorText.Visibility = Visibility.Visible;
                return;
            }

            string conflictSubject;
            if (HasConflictingPaymentSelections(out conflictSubject))
            {
                ErrorText.Text = "同一主体出现多个支付方选择：" + conflictSubject + "。请重新预检。";
                ErrorText.Visibility = Visibility.Visible;
                return;
            }

            if (HasConflictingTemplateSelections(out conflictSubject))
            {
                ErrorText.Text = "同一主体出现多个分表模板选择：" + conflictSubject + "。请重新预检。";
                ErrorText.Visibility = Visibility.Visible;
                return;
            }

            PaymentDecisions = BuildPaymentDecisions();
            TemplateDecisions = BuildTemplateDecisions();
            DialogResult = true;
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }

        private void Decision_SelectionChanged(object sender, SelectionChangedEventArgs e)
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
                    .All(row => (!row.RequiresPaymentPartySelection || HasValidPaymentSelection(row))
                        && (!row.RequiresTemplateSelection || HasValidTemplateSelection(row)));
        }

        private static string BuildSubjectText(Stage2PreflightIssueItemViewModel row)
        {
            var parts = new[] { EffectiveSettlementKind(row), row.Entity }
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

        private static string BuildMissingDecisionMessage(
            IList<string> missingPayments,
            IList<string> missingTemplates)
        {
            var parts = new List<string>();
            if (missingPayments.Count > 0)
            {
                parts.Add("支付方：" + BuildShortSubjectList(missingPayments));
            }

            if (missingTemplates.Count > 0)
            {
                parts.Add("分表模板：" + BuildShortSubjectList(missingTemplates));
            }

            return "请先完成以下必选项：" + string.Join("；", parts) + "。";
        }

        private List<Stage2PreflightPaymentDecision> BuildPaymentDecisions()
        {
            return _viewModel.IssueRows
                .Where(row => row.RequiresPaymentPartySelection)
                .GroupBy(DecisionKey, StringComparer.Ordinal)
                .Select(group => group.First())
                .Select(row => new Stage2PreflightPaymentDecision
                {
                    SettlementKind = EffectiveSettlementKind(row),
                    Entity = row.Entity,
                    PaymentParty = row.SelectedPaymentParty
                })
                .ToList();
        }

        private List<Stage2PreflightTemplateDecision> BuildTemplateDecisions()
        {
            return _viewModel.IssueRows
                .Where(row => row.RequiresTemplateSelection)
                .GroupBy(DecisionKey, StringComparer.Ordinal)
                .Select(group => group.First())
                .Select(row => new Stage2PreflightTemplateDecision
                {
                    SettlementKind = EffectiveSettlementKind(row),
                    Entity = row.Entity,
                    TemplatePath = row.SelectedTemplatePath
                })
                .ToList();
        }

        private bool HasConflictingPaymentSelections(out string subject)
        {
            var conflict = _viewModel.IssueRows
                .Where(row => row.RequiresPaymentPartySelection)
                .GroupBy(DecisionKey, StringComparer.Ordinal)
                .FirstOrDefault(group => group
                    .Select(row => row.SelectedPaymentParty)
                    .Where(value => !string.IsNullOrWhiteSpace(value))
                    .Distinct(StringComparer.Ordinal)
                    .Skip(1)
                    .Any());
            subject = conflict == null ? null : BuildSubjectText(conflict.First());
            return conflict != null;
        }

        private bool HasConflictingTemplateSelections(out string subject)
        {
            var conflict = _viewModel.IssueRows
                .Where(row => row.RequiresTemplateSelection)
                .GroupBy(DecisionKey, StringComparer.Ordinal)
                .FirstOrDefault(group => group
                    .Select(row => row.SelectedTemplatePath)
                    .Where(value => !string.IsNullOrWhiteSpace(value))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .Skip(1)
                    .Any());
            subject = conflict == null ? null : BuildSubjectText(conflict.First());
            return conflict != null;
        }

        private static bool HasValidPaymentSelection(Stage2PreflightIssueItemViewModel row)
        {
            return !string.IsNullOrWhiteSpace(row.SelectedPaymentParty)
                && row.PaymentPartyOptions != null
                && row.PaymentPartyOptions.Contains(row.SelectedPaymentParty, StringComparer.Ordinal);
        }

        private static bool HasValidTemplateSelection(Stage2PreflightIssueItemViewModel row)
        {
            return !string.IsNullOrWhiteSpace(row.SelectedTemplatePath)
                && row.TemplateOptions != null
                && row.TemplateOptions.Any(option => string.Equals(
                    option.Path,
                    row.SelectedTemplatePath,
                    StringComparison.OrdinalIgnoreCase));
        }

        private static string DecisionKey(Stage2PreflightIssueItemViewModel row)
        {
            return TextUtil.S(EffectiveSettlementKind(row))
                + "\u001f"
                + TextUtil.CustomerKey(row.Entity);
        }

        private static string EffectiveSettlementKind(Stage2PreflightIssueItemViewModel row)
        {
            return string.IsNullOrWhiteSpace(row.SettlementKind) ? row.Kind : row.SettlementKind;
        }
    }
}
