using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using HainanSettlementTool.Core.Models;
using HainanSettlementTool.Core.Services;

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
                    IntroText = "每个主体已集中显示需要你完成的选择和程序将自动填写的内容。请先处理红色阻断项，并完成所有必选项。",
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
                    IntroText = "每个主体已集中显示需要你完成的选择和程序将自动填写的内容。请先处理红色阻断项，并完成所有必选项。",
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
                .GroupBy(BuildCardKey)
                .Select(BuildIssueGroup)
                .OrderBy(group => DispositionOrder(group.Disposition))
                .ThenBy(group => group.Heading, StringComparer.CurrentCulture)
                .ToList();
            foreach (var group in groups)
            {
                group.CountText = group.Issues.Count + " 项";
            }

            var effectiveBlockerCount = groups.Count(group =>
                group.Disposition == Stage2PreflightDisposition.Blocker);
            var requiredDecisionCount = groups.Count(group =>
                group.Disposition == Stage2PreflightDisposition.RequiredDecision);
            var reviewCount = groups.Count(group =>
                group.Disposition == Stage2PreflightDisposition.Review);
            var informationCount = groups.Count(group =>
                group.Disposition == Stage2PreflightDisposition.Information);
            var subjectCardCount = groups.Count(group => group.IsSubjectGroup);
            var generalCardCount = groups.Count - subjectCardCount;
            return new Stage2PreflightDialogViewModel
            {
                Title = profile.Title,
                Heading = profile.Heading,
                IntroText = profile.IntroText,
                ConfirmButtonText = profile.ConfirmButtonText,
                SummaryText = BuildSummaryText(
                    month,
                    subjectCount,
                    subjectCardCount,
                    generalCardCount,
                    issueRows.Count),
                BlockerCountText = "阻断 " + effectiveBlockerCount,
                RequiredDecisionCountText = "必选主体 " + requiredDecisionCount,
                ReviewCountText = "仅复核 " + reviewCount,
                InformationCountText = "仅信息 " + informationCount,
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
            var issueDecisionKey = DecisionKey(settlementKind, issue.Entity);
            var resolvedDecision = evaluation.DecisionResolutions.FirstOrDefault(item =>
                item.Status == Stage2PaymentPartyDecisionStatus.Resolved
                && DecisionKey(item.SettlementKind, item.Entity) == issueDecisionKey);
            var resolvedTemplateDecision = evaluation.TemplateDecisionResolutions.FirstOrDefault(item =>
                item.Status == Stage2TemplateDecisionStatus.Resolved
                && DecisionKey(item.SettlementKind, item.Entity) == issueDecisionKey);
            var category = DisplayCategory(issue);
            var contextText = BuildContextText(issue);
            var technicalDetailsText = BuildTechnicalDetailsText(issue);
            var templateOptions = BuildTemplateOptions(issue);
            var row = new Stage2PreflightIssueItemViewModel
            {
                Category = category,
                Disposition = issue.Disposition,
                StatusText = DispositionText(issue.Disposition),
                PrimaryText = string.IsNullOrWhiteSpace(issue.Message) ? category : issue.Message,
                Kind = issue.Kind,
                HandlingText = BuildHandlingText(issue),
                ContextText = contextText,
                ContextTextVisibility = string.IsNullOrWhiteSpace(contextText) ? Visibility.Collapsed : Visibility.Visible,
                TechnicalDetailsText = technicalDetailsText,
                TechnicalDetailsVisibility = string.IsNullOrWhiteSpace(technicalDetailsText)
                    ? Visibility.Collapsed
                    : Visibility.Visible,
                SettlementKind = settlementKind,
                Entity = issue.Entity,
                Owner = issue.Owner,
                LedgerRow = issue.LedgerRow,
                RequiresPaymentPartySelection = issue.RequiresPaymentPartySelection,
                PaymentSelectionVisibility = issue.RequiresPaymentPartySelection
                    ? Visibility.Visible
                    : Visibility.Collapsed,
                PaymentPartyOptions = (issue.PaymentPartyOptions ?? new string[0]).ToList(),
                SelectedPaymentParty = resolvedDecision == null ? null : resolvedDecision.PaymentParty,
                RequiresTemplateSelection = issue.RequiresTemplateSelection,
                TemplateSelectionVisibility = issue.RequiresTemplateSelection
                    ? Visibility.Visible
                    : Visibility.Collapsed,
                TemplateOptions = templateOptions
            };
            if (issue.RequiresTemplateSelection)
            {
                row.TemplateBrowser = new Stage2TemplateCandidateBrowserViewModel(
                    templateOptions,
                    resolvedTemplateDecision == null ? null : resolvedTemplateDecision.TemplatePath);
            }

            return row;
        }

        private static string DisplayCategory(Stage2PreflightIssue issue)
        {
            if (issue.RequiresTemplateSelection)
            {
                return "需要你选择：新分表模板";
            }

            if (issue.RequiresPaymentPartySelection)
            {
                return "需要你选择：支付方";
            }

            if (issue.Code == Stage2PreflightIssueKinds.NewSummarySubject)
            {
                return "程序将填写：新增主体默认资料";
            }

            if (issue.Code == Stage2PreflightIssueKinds.BorrowedTemplate)
            {
                return "程序将使用：同类型分表模板";
            }

            return string.IsNullOrWhiteSpace(issue.Category) ? "其他预检项目" : issue.Category;
        }

        private static string BuildCardKey(Stage2PreflightIssueItemViewModel row)
        {
            if (HasSubjectIdentity(row))
            {
                return "subject\u001f" + DecisionKey(EffectiveSettlementKind(row), row.Entity);
            }

            return "category\u001f" + row.Category;
        }

        private static string DecisionKey(string settlementKind, string entity)
        {
            return TextUtil.S(settlementKind) + "\u001f" + TextUtil.CustomerKey(entity);
        }

        private static Stage2PreflightIssueGroupViewModel BuildIssueGroup(
            IGrouping<string, Stage2PreflightIssueItemViewModel> source)
        {
            var issues = source
                .OrderBy(row => DispositionOrder(row.Disposition))
                .ThenBy(row => row.Category, StringComparer.CurrentCulture)
                .ThenBy(row => row.PrimaryText, StringComparer.CurrentCulture)
                .ToList();
            var first = issues[0];
            var isSubjectGroup = HasSubjectIdentity(first);
            var disposition = issues
                .OrderBy(row => DispositionOrder(row.Disposition))
                .First()
                .Disposition;
            return new Stage2PreflightIssueGroupViewModel
            {
                Heading = isSubjectGroup
                    ? EffectiveSettlementKind(first) + " · " + first.Entity
                    : first.Category,
                SupportingText = isSubjectGroup ? BuildGroupSupportingText(issues) : string.Empty,
                SupportingTextVisibility = isSubjectGroup ? Visibility.Visible : Visibility.Collapsed,
                IsSubjectGroup = isSubjectGroup,
                Disposition = disposition,
                StatusText = DispositionText(disposition),
                Issues = issues
            };
        }

        private static string BuildGroupSupportingText(
            IEnumerable<Stage2PreflightIssueItemViewModel> issues)
        {
            var rows = issues.ToList();
            var owners = rows
                .Select(row => row.Owner)
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Distinct(StringComparer.CurrentCulture)
                .ToList();
            var ledgerRows = rows
                .Select(row => row.LedgerRow)
                .Where(value => value > 0)
                .Distinct()
                .OrderBy(value => value)
                .ToList();
            var parts = new List<string>();
            if (owners.Count > 0)
            {
                parts.Add("负责人：" + string.Join("、", owners));
            }

            if (ledgerRows.Count > 0)
            {
                parts.Add("台账行：" + string.Join("、", ledgerRows));
            }

            return string.Join("   ", parts);
        }

        private static bool HasSubjectIdentity(Stage2PreflightIssueItemViewModel row)
        {
            return !string.IsNullOrWhiteSpace(row.Entity)
                && !string.IsNullOrWhiteSpace(EffectiveSettlementKind(row));
        }

        private static string EffectiveSettlementKind(Stage2PreflightIssueItemViewModel row)
        {
            return string.IsNullOrWhiteSpace(row.SettlementKind) ? row.Kind : row.SettlementKind;
        }

        private static string BuildSummaryText(
            int month,
            int subjectCount,
            int subjectCardCount,
            int generalCardCount,
            int issueCount)
        {
            var scope = subjectCardCount + " 个主体";
            if (generalCardCount > 0)
            {
                scope += "、" + generalCardCount + " 个通用问题组";
            }

            return "结算月份：2026年" + month + "月；本月应生成主体 " + subjectCount
                + " 个；预检涉及 " + scope + "，共 " + issueCount + " 项检查。";
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
            rows.AddRange(evaluation.TemplateDecisionResolutions
                .Where(item => item.Status == Stage2TemplateDecisionStatus.Invalid
                    || item.Status == Stage2TemplateDecisionStatus.Conflicting
                    || item.Status == Stage2TemplateDecisionStatus.Stale)
                .Select(item => BuildDiagnosticRow(
                    "分表模板选择异常",
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
                HandlingText = "处理方式：请修正输入或预检定义后重新预检；程序不会在当前状态下继续生成。",
                ContextText = string.Join("；", context),
                ContextTextVisibility = context.Count > 0 ? Visibility.Visible : Visibility.Collapsed,
                TechnicalDetailsVisibility = Visibility.Collapsed,
                SettlementKind = settlementKind,
                Entity = entity,
                PaymentSelectionVisibility = Visibility.Collapsed,
                PaymentPartyOptions = new List<string>(),
                TemplateSelectionVisibility = Visibility.Collapsed,
                TemplateOptions = new List<Stage2PreflightTemplateOptionViewModel>()
            };
        }

        private static List<Stage2PreflightTemplateOptionViewModel> BuildTemplateOptions(
            Stage2PreflightIssue issue)
        {
            IEnumerable<string> paths = issue.TemplateOptions == null
                ? Enumerable.Empty<string>()
                : (IEnumerable<string>)issue.TemplateOptions;
            return paths
                .Where(path => !string.IsNullOrWhiteSpace(path))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Select(BuildTemplateOption)
                .OrderBy(option => option.DisplayText, StringComparer.CurrentCulture)
                .ToList();
        }

        private static Stage2PreflightTemplateOptionViewModel BuildTemplateOption(string path)
        {
            var fileName = SafeFileName(path);
            var subject = NormalizeTemplateSubject(
                Path.GetFileNameWithoutExtension(fileName) ?? string.Empty);

            var owner = string.Empty;
            try
            {
                var directory = Path.GetDirectoryName(path);
                owner = string.IsNullOrWhiteSpace(directory)
                    ? string.Empty
                    : new DirectoryInfo(directory).Name;
            }
            catch
            {
                owner = string.Empty;
            }
            owner = NormalizeTemplateOwner(owner);

            var parts = new[] { subject, owner, fileName }
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Distinct(StringComparer.CurrentCulture)
                .ToList();
            return new Stage2PreflightTemplateOptionViewModel
            {
                Path = path,
                SubjectText = string.IsNullOrWhiteSpace(subject) ? fileName : subject,
                OwnerText = owner,
                FileName = fileName,
                DisplayText = parts.Count == 0 ? fileName : string.Join(" / ", parts)
            };
        }

        private static string NormalizeTemplateSubject(string subject)
        {
            var value = (subject ?? string.Empty).Trim();
            var yearMarker = value.LastIndexOf(" 20", StringComparison.Ordinal);
            if (yearMarker <= 0)
            {
                return value;
            }

            var suffix = value.Substring(yearMarker + 1);
            return suffix.Length > 4 && ContainsFourDigitYear(suffix)
                ? value.Substring(0, yearMarker).Trim()
                : value;
        }

        private static string NormalizeTemplateOwner(string owner)
        {
            var value = (owner ?? string.Empty).Trim();
            var separator = value.LastIndexOf(" - ", StringComparison.Ordinal);
            if (separator <= 0)
            {
                return value;
            }

            var suffix = value.Substring(separator + 3).Trim();
            return ContainsFourDigitYear(suffix)
                ? value.Substring(0, separator).Trim()
                : value;
        }

        private static bool ContainsFourDigitYear(string value)
        {
            if (string.IsNullOrWhiteSpace(value) || value.Length < 4)
            {
                return false;
            }

            for (var index = 0; index <= value.Length - 4; index++)
            {
                if (!char.IsDigit(value[index])
                    || !char.IsDigit(value[index + 1])
                    || !char.IsDigit(value[index + 2])
                    || !char.IsDigit(value[index + 3]))
                {
                    continue;
                }

                var hasDigitBefore = index > 0 && char.IsDigit(value[index - 1]);
                var hasDigitAfter = index + 4 < value.Length && char.IsDigit(value[index + 4]);
                if (!hasDigitBefore && !hasDigitAfter)
                {
                    return true;
                }
            }

            return false;
        }

        private static string SafeFileName(string path)
        {
            try
            {
                var fileName = Path.GetFileName(path);
                return string.IsNullOrWhiteSpace(fileName) ? path : fileName;
            }
            catch
            {
                return path;
            }
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
            if (!string.IsNullOrWhiteSpace(issue.Kind)
                && !string.IsNullOrWhiteSpace(issue.SettlementKind)
                && !string.Equals(issue.Kind, issue.SettlementKind, StringComparison.Ordinal))
            {
                parts.Add("问题：" + issue.Kind);
            }

            if (!string.IsNullOrWhiteSpace(issue.Customer))
            {
                parts.Add("客户/明细：" + issue.Customer);
            }

            if ((!string.IsNullOrWhiteSpace(issue.PreviousValue)
                    && !LooksLikeFilePathList(issue.PreviousValue))
                || (!string.IsNullOrWhiteSpace(issue.CurrentValue)
                    && !LooksLikeFilePathList(issue.CurrentValue)))
            {
                if (!string.IsNullOrWhiteSpace(issue.PreviousValue)
                    && !LooksLikeFilePathList(issue.PreviousValue))
                {
                    parts.Add("原值：" + issue.PreviousValue);
                }

                if (!string.IsNullOrWhiteSpace(issue.CurrentValue)
                    && !LooksLikeFilePathList(issue.CurrentValue))
                {
                    parts.Add("当前/将写入：" + issue.CurrentValue);
                }
            }

            if (!string.IsNullOrWhiteSpace(issue.SheetName))
            {
                parts.Add("工作表：" + issue.SheetName);
            }

            return string.Join("；", parts);
        }

        private static string BuildTechnicalDetailsText(Stage2PreflightIssue issue)
        {
            var parts = new List<string>();
            if (!string.IsNullOrWhiteSpace(issue.TemplateFile))
            {
                parts.Add("文件名：" + SafeFileName(issue.TemplateFile));
                parts.Add("完整路径：" + issue.TemplateFile);
            }

            IEnumerable<string> templateOptions = issue.TemplateOptions == null
                ? Enumerable.Empty<string>()
                : (IEnumerable<string>)issue.TemplateOptions;
            var candidates = templateOptions
                .Where(path => !string.IsNullOrWhiteSpace(path))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            if (candidates.Count == 0 && LooksLikeFilePathList(issue.CurrentValue))
            {
                candidates = SplitPathList(issue.CurrentValue);
            }

            if (candidates.Count > 0)
            {
                parts.Add("候选摘要：共 " + candidates.Count + " 个工作簿，仅用于选择新分表的版式来源。");
                parts.Add("候选文件：" + string.Join("、", candidates.Select(SafeFileName)));
                parts.Add("候选完整路径：" + Environment.NewLine
                    + string.Join(Environment.NewLine, candidates));
            }
            else
            {
                if (LooksLikeFilePathList(issue.PreviousValue))
                {
                    parts.Add("原始路径：" + issue.PreviousValue);
                }

                if (LooksLikeFilePathList(issue.CurrentValue))
                {
                    parts.Add("当前路径：" + issue.CurrentValue);
                }
            }

            return string.Join(Environment.NewLine, parts);
        }

        private static bool LooksLikeFilePathList(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            return value.IndexOf(":\\", StringComparison.Ordinal) >= 0
                || value.StartsWith("\\\\", StringComparison.Ordinal)
                || value.IndexOf(".xlsx", StringComparison.OrdinalIgnoreCase) >= 0
                || value.IndexOf(".xls", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static List<string> SplitPathList(string value)
        {
            return (value ?? string.Empty)
                .Split(new[] { '、', '；', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(item => item.Trim())
                .Where(item => !string.IsNullOrWhiteSpace(item))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
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
        public string Heading { get; set; }
        public string SupportingText { get; set; }
        public Visibility SupportingTextVisibility { get; set; }
        public bool IsSubjectGroup { get; set; }
        public Stage2PreflightDisposition Disposition { get; set; }
        public string StatusText { get; set; }
        public string CountText { get; set; }
        public List<Stage2PreflightIssueItemViewModel> Issues { get; set; }
    }

    internal sealed class Stage2PreflightIssueItemViewModel
    {
        public string Category { get; set; }
        public Stage2PreflightDisposition Disposition { get; set; }
        public string StatusText { get; set; }
        public string PrimaryText { get; set; }
        public string Kind { get; set; }
        public string HandlingText { get; set; }
        public string ContextText { get; set; }
        public Visibility ContextTextVisibility { get; set; }
        public string TechnicalDetailsText { get; set; }
        public Visibility TechnicalDetailsVisibility { get; set; }
        public string SettlementKind { get; set; }
        public string Entity { get; set; }
        public string Owner { get; set; }
        public int LedgerRow { get; set; }
        public bool RequiresPaymentPartySelection { get; set; }
        public Visibility PaymentSelectionVisibility { get; set; }
        public List<string> PaymentPartyOptions { get; set; }
        public string SelectedPaymentParty { get; set; }
        public bool RequiresTemplateSelection { get; set; }
        public Visibility TemplateSelectionVisibility { get; set; }
        public List<Stage2PreflightTemplateOptionViewModel> TemplateOptions { get; set; }
        public Stage2TemplateCandidateBrowserViewModel TemplateBrowser { get; set; }
        public string SelectedTemplatePath
        {
            get { return TemplateBrowser == null ? null : TemplateBrowser.SelectedTemplatePath; }
        }
    }

    internal sealed class Stage2PreflightTemplateOptionViewModel : System.ComponentModel.INotifyPropertyChanged
    {
        private bool _isSelected;

        public event System.ComponentModel.PropertyChangedEventHandler PropertyChanged;

        public string DisplayText { get; set; }
        public string Path { get; set; }
        public string SubjectText { get; set; }
        public string OwnerText { get; set; }
        public string FileName { get; set; }
        public bool IsSelected
        {
            get { return _isSelected; }
            set
            {
                if (_isSelected == value)
                {
                    return;
                }

                _isSelected = value;
                PropertyChanged?.Invoke(
                    this,
                    new System.ComponentModel.PropertyChangedEventArgs(nameof(IsSelected)));
            }
        }
    }

    internal sealed class Stage2PreflightPaymentDecision
    {
        public string SettlementKind { get; set; }
        public string Entity { get; set; }
        public string PaymentParty { get; set; }
    }

    internal sealed class Stage2PreflightTemplateDecision
    {
        public string SettlementKind { get; set; }
        public string Entity { get; set; }
        public string TemplatePath { get; set; }
    }
}
