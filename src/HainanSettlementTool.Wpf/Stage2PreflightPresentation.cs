using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using HainanSettlementTool.Core.Models;

namespace HainanSettlementTool.Wpf
{
    internal static class Stage2PreflightPresentationAdapter
    {
        public static Stage2PreflightDialogViewModel CreateHainan(
            HainanStage2PreflightReport report,
            Stage2PreflightEvaluation evaluation)
        {
            if (report == null)
            {
                throw new ArgumentNullException(nameof(report));
            }

            return Create(
                new Stage2PreflightDialogProfile
                {
                    Title = "海南阶段二预检确认",
                    Heading = "海南阶段二预检",
                    IntroText = "请先处理红色阻断项；新增主体或存量支付方缺失时必须完成选择。复核和信息项会随本次生成写入校验报告。",
                    ConfirmButtonText = "继续生成并写报告"
                },
                report.Month,
                report.SubjectCount,
                report.Issues.Cast<Stage2PreflightIssue>(),
                evaluation);
        }

        public static Stage2PreflightDialogViewModel CreateChongqing(
            ChongqingStage2PreflightReport report,
            Stage2PreflightEvaluation evaluation)
        {
            if (report == null)
            {
                throw new ArgumentNullException(nameof(report));
            }

            return Create(
                new Stage2PreflightDialogProfile
                {
                    Title = "重庆阶段二预检确认",
                    Heading = "重庆阶段二预检",
                    IntroText = "请先处理红色阻断项；新增主体或存量支付方缺失时必须完成选择。复核和信息项会随本次生成写入校验报告。",
                    ConfirmButtonText = "确认并生成"
                },
                report.Month,
                report.SubjectCount,
                report.Issues.Cast<Stage2PreflightIssue>(),
                evaluation);
        }

        private static Stage2PreflightDialogViewModel Create(
            Stage2PreflightDialogProfile profile,
            int month,
            int subjectCount,
            IEnumerable<Stage2PreflightIssue> issues,
            Stage2PreflightEvaluation evaluation)
        {
            if (evaluation == null)
            {
                throw new ArgumentNullException(nameof(evaluation));
            }

            var issueRows = (issues ?? Enumerable.Empty<Stage2PreflightIssue>())
                .Where(issue => issue != null)
                .Select(issue => BuildIssueRow(issue, evaluation))
                .ToList();
            var diagnosticRows = BuildBlockingDiagnosticRows(evaluation);
            issueRows.AddRange(diagnosticRows);

            var groups = issueRows
                .GroupBy(row => row.Category)
                .Select(group => new Stage2PreflightIssueGroupViewModel
                {
                    Category = group.Key,
                    Issues = group
                        .OrderBy(row => DispositionOrder(row.Disposition))
                        .ThenBy(row => row.PrimaryText, StringComparer.CurrentCulture)
                        .ToList()
                })
                .OrderBy(group => group.Issues.Min(row => DispositionOrder(row.Disposition)))
                .ThenBy(group => group.Category, StringComparer.CurrentCulture)
                .ToList();
            foreach (var group in groups)
            {
                group.CountText = group.Issues.Count + " 条";
            }

            var effectiveBlockerCount = issueRows.Count(row =>
                row.Disposition == Stage2PreflightDisposition.Blocker);
            var requiredDecisionCount = issueRows.Count(row =>
                row.Disposition == Stage2PreflightDisposition.RequiredDecision);
            var reviewCount = issueRows.Count(row =>
                row.Disposition == Stage2PreflightDisposition.Review);
            var informationCount = issueRows.Count(row =>
                row.Disposition == Stage2PreflightDisposition.Information);
            return new Stage2PreflightDialogViewModel
            {
                Title = profile.Title,
                Heading = profile.Heading,
                IntroText = profile.IntroText,
                ConfirmButtonText = profile.ConfirmButtonText,
                SummaryText = "结算月份：2026年" + month + "月；本月应生成主体 " + subjectCount
                    + " 个；共发现 " + issueRows.Count + " 条预检项目。",
                BlockerCountText = "阻断 " + effectiveBlockerCount,
                RequiredDecisionCountText = "必选 " + requiredDecisionCount,
                ReviewCountText = "复核 " + reviewCount,
                InformationCountText = "信息 " + informationCount,
                BlockingMessage = effectiveBlockerCount > 0
                    ? "存在 " + effectiveBlockerCount + " 个阻断项。请按红色项目提示修正后重新预检，当前不能继续生成。"
                    : string.Empty,
                BlockingMessageVisibility = effectiveBlockerCount > 0 ? Visibility.Visible : Visibility.Collapsed,
                CanConfirm = effectiveBlockerCount == 0 && !evaluation.HasInvalidDecisions,
                IssueGroups = groups,
                IssueRows = issueRows
            };
        }

        private static Stage2PreflightIssueItemViewModel BuildIssueRow(
            Stage2PreflightIssue issue,
            Stage2PreflightEvaluation evaluation)
        {
            var settlementKind = issue.SettlementKind;
            var resolvedDecision = evaluation.DecisionResolutions.FirstOrDefault(item =>
                item.Status == Stage2PaymentPartyDecisionStatus.Resolved
                && string.Equals(item.SettlementKind, settlementKind, StringComparison.Ordinal)
                && string.Equals(item.Entity, issue.Entity, StringComparison.Ordinal));
            var category = string.IsNullOrWhiteSpace(issue.Category) ? "其他预检项目" : issue.Category;
            var contextText = BuildContextText(issue);
            return new Stage2PreflightIssueItemViewModel
            {
                Category = category,
                Disposition = issue.Disposition,
                StatusText = DispositionText(issue.Disposition),
                PrimaryText = string.IsNullOrWhiteSpace(issue.Message) ? category : issue.Message,
                FileText = BuildFileText(issue),
                FileTextVisibility = string.IsNullOrWhiteSpace(issue.TemplateFile)
                    ? Visibility.Collapsed
                    : Visibility.Visible,
                HandlingText = BuildHandlingText(issue),
                ContextText = contextText,
                ContextTextVisibility = string.IsNullOrWhiteSpace(contextText) ? Visibility.Collapsed : Visibility.Visible,
                SettlementKind = settlementKind,
                Entity = issue.Entity,
                RequiresPaymentPartySelection = issue.RequiresPaymentPartySelection,
                PaymentSelectionVisibility = issue.RequiresPaymentPartySelection
                    ? Visibility.Visible
                    : Visibility.Collapsed,
                PaymentPartyOptions = (issue.PaymentPartyOptions ?? new string[0]).ToList(),
                SelectedPaymentParty = resolvedDecision == null ? null : resolvedDecision.PaymentParty
            };
        }

        private static List<Stage2PreflightIssueItemViewModel> BuildBlockingDiagnosticRows(
            Stage2PreflightEvaluation evaluation)
        {
            var rows = evaluation.InvalidDefinitions
                .Where(message => !string.IsNullOrWhiteSpace(message))
                .Select(message => BuildDiagnosticRow("预检定义异常", message, null, null))
                .ToList();
            rows.AddRange(evaluation.DecisionResolutions
                .Where(item => item.Status == Stage2PaymentPartyDecisionStatus.Invalid
                    || item.Status == Stage2PaymentPartyDecisionStatus.Conflicting
                    || item.Status == Stage2PaymentPartyDecisionStatus.Stale)
                .Select(item => BuildDiagnosticRow(
                    "支付方选择异常",
                    item.Message,
                    item.SettlementKind,
                    item.Entity)));
            return rows;
        }

        private static Stage2PreflightIssueItemViewModel BuildDiagnosticRow(
            string category,
            string message,
            string settlementKind,
            string entity)
        {
            var context = new List<string>();
            if (!string.IsNullOrWhiteSpace(settlementKind))
            {
                context.Add("结算类型：" + settlementKind);
            }

            if (!string.IsNullOrWhiteSpace(entity))
            {
                context.Add("主体：" + entity);
            }

            return new Stage2PreflightIssueItemViewModel
            {
                Category = category,
                Disposition = Stage2PreflightDisposition.Blocker,
                StatusText = "阻断",
                PrimaryText = string.IsNullOrWhiteSpace(message) ? "阶段二预检数据不完整。" : message,
                FileText = string.Empty,
                FileTextVisibility = Visibility.Collapsed,
                HandlingText = "处理方式：请修正输入或预检定义后重新预检；程序不会在当前状态下继续生成。",
                ContextText = string.Join("；", context),
                ContextTextVisibility = context.Count > 0 ? Visibility.Visible : Visibility.Collapsed,
                SettlementKind = settlementKind,
                Entity = entity,
                PaymentSelectionVisibility = Visibility.Collapsed,
                PaymentPartyOptions = new List<string>()
            };
        }

        private static string BuildFileText(Stage2PreflightIssue issue)
        {
            if (string.IsNullOrWhiteSpace(issue.TemplateFile))
            {
                return string.Empty;
            }

            return "相关文件：" + Path.GetFileName(issue.TemplateFile);
        }

        private static string BuildHandlingText(Stage2PreflightIssue issue)
        {
            return string.IsNullOrWhiteSpace(issue.Suggestion)
                ? "处理方式：请根据预检项目说明人工确认。"
                : "处理方式：" + issue.Suggestion;
        }

        private static string BuildContextText(Stage2PreflightIssue issue)
        {
            var parts = new List<string>();
            if (!string.IsNullOrWhiteSpace(issue.SettlementKind))
            {
                parts.Add("结算类型：" + issue.SettlementKind);
            }
            else if (!string.IsNullOrWhiteSpace(issue.Kind))
            {
                parts.Add("类型：" + issue.Kind);
            }

            if (!string.IsNullOrWhiteSpace(issue.Kind)
                && !string.Equals(issue.Kind, issue.SettlementKind, StringComparison.Ordinal))
            {
                parts.Add("问题：" + issue.Kind);
            }

            if (!string.IsNullOrWhiteSpace(issue.Owner) || !string.IsNullOrWhiteSpace(issue.Entity))
            {
                parts.Add("负责人/主体：" + TextOrDash(issue.Owner) + " / " + TextOrDash(issue.Entity));
            }

            if (!string.IsNullOrWhiteSpace(issue.Customer))
            {
                parts.Add("客户/明细：" + issue.Customer);
            }

            if (issue.LedgerRow > 0)
            {
                parts.Add("台账行：" + issue.LedgerRow);
            }

            if (!string.IsNullOrWhiteSpace(issue.PreviousValue) || !string.IsNullOrWhiteSpace(issue.CurrentValue))
            {
                parts.Add("对比：" + TextOrDash(issue.PreviousValue) + "；" + TextOrDash(issue.CurrentValue));
            }

            if (!string.IsNullOrWhiteSpace(issue.SheetName))
            {
                parts.Add("工作表：" + issue.SheetName);
            }

            if (!string.IsNullOrWhiteSpace(issue.TemplateFile))
            {
                parts.Add("完整路径：" + issue.TemplateFile);
            }

            return string.Join("；", parts);
        }

        private static string TextOrDash(string value)
        {
            return string.IsNullOrWhiteSpace(value) ? "-" : value;
        }

        private static string DispositionText(Stage2PreflightDisposition disposition)
        {
            switch (disposition)
            {
                case Stage2PreflightDisposition.Blocker:
                    return "阻断";
                case Stage2PreflightDisposition.RequiredDecision:
                    return "必选";
                case Stage2PreflightDisposition.Information:
                    return "信息";
                default:
                    return "复核";
            }
        }

        private static int DispositionOrder(Stage2PreflightDisposition disposition)
        {
            switch (disposition)
            {
                case Stage2PreflightDisposition.Blocker:
                    return 0;
                case Stage2PreflightDisposition.RequiredDecision:
                    return 1;
                case Stage2PreflightDisposition.Review:
                    return 2;
                default:
                    return 3;
            }
        }

        private sealed class Stage2PreflightDialogProfile
        {
            public string Title { get; set; }
            public string Heading { get; set; }
            public string IntroText { get; set; }
            public string ConfirmButtonText { get; set; }
        }
    }

    internal sealed class Stage2PreflightDialogViewModel
    {
        public string Title { get; set; }
        public string Heading { get; set; }
        public string IntroText { get; set; }
        public string ConfirmButtonText { get; set; }
        public string SummaryText { get; set; }
        public string BlockerCountText { get; set; }
        public string RequiredDecisionCountText { get; set; }
        public string ReviewCountText { get; set; }
        public string InformationCountText { get; set; }
        public string BlockingMessage { get; set; }
        public Visibility BlockingMessageVisibility { get; set; }
        public bool CanConfirm { get; set; }
        public List<Stage2PreflightIssueGroupViewModel> IssueGroups { get; set; }
        public List<Stage2PreflightIssueItemViewModel> IssueRows { get; set; }
    }

    internal sealed class Stage2PreflightIssueGroupViewModel
    {
        public string Category { get; set; }
        public string CountText { get; set; }
        public List<Stage2PreflightIssueItemViewModel> Issues { get; set; }
    }

    internal sealed class Stage2PreflightIssueItemViewModel
    {
        public string Category { get; set; }
        public Stage2PreflightDisposition Disposition { get; set; }
        public string StatusText { get; set; }
        public string PrimaryText { get; set; }
        public string FileText { get; set; }
        public Visibility FileTextVisibility { get; set; }
        public string HandlingText { get; set; }
        public string ContextText { get; set; }
        public Visibility ContextTextVisibility { get; set; }
        public string SettlementKind { get; set; }
        public string Entity { get; set; }
        public bool RequiresPaymentPartySelection { get; set; }
        public Visibility PaymentSelectionVisibility { get; set; }
        public List<string> PaymentPartyOptions { get; set; }
        public string SelectedPaymentParty { get; set; }
    }

    internal sealed class Stage2PreflightPaymentDecision
    {
        public string SettlementKind { get; set; }
        public string Entity { get; set; }
        public string PaymentParty { get; set; }
    }
}
