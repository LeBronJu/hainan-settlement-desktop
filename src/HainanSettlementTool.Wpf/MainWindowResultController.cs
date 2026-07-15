using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using HainanSettlementTool.Core.Models;

namespace HainanSettlementTool.Wpf
{
    internal sealed class MainWindowResultController
    {
        private readonly Border _completionCard;
        private readonly Border _completionIconCircle;
        private readonly TextBlock _completionIconText;
        private readonly TextBlock _completionTitleText;
        private readonly TextBlock _completionDetailText;
        private readonly TextBlock _completionOutputLabel;
        private readonly TextBlock _completionOutputText;
        private readonly TextBlock _noProvinceResultHint;
        private readonly Grid _stage1ResultRow;
        private readonly TextBlock _stage1ResultLabel;
        private readonly TextBlock _stage1ResultStatus;
        private readonly TextBlock _stage1ResultCount;
        private readonly Grid _proxyResultRow;
        private readonly TextBlock _proxyResultStatus;
        private readonly TextBlock _proxyResultCount;
        private readonly Grid _intermediaryResultRow;
        private readonly TextBlock _intermediaryResultStatus;
        private readonly TextBlock _intermediaryResultCount;
        private readonly Grid _summaryResultRow;
        private readonly TextBlock _summaryResultLabel;
        private readonly TextBlock _summaryResultStatus;
        private readonly TextBlock _summaryResultCount;
        private readonly Grid _employeeRewardResultRow;
        private readonly TextBlock _employeeRewardResultStatus;
        private readonly TextBlock _employeeRewardResultCount;
        private readonly DockPanel _finishedAtRow;
        private readonly TextBlock _finishedAtText;
        private readonly Func<string, Brush> _brushOf;

        public MainWindowResultController(
            Border completionCard,
            Border completionIconCircle,
            TextBlock completionIconText,
            TextBlock completionTitleText,
            TextBlock completionDetailText,
            TextBlock completionOutputLabel,
            TextBlock completionOutputText,
            TextBlock noProvinceResultHint,
            Grid stage1ResultRow,
            TextBlock stage1ResultLabel,
            TextBlock stage1ResultStatus,
            TextBlock stage1ResultCount,
            Grid proxyResultRow,
            TextBlock proxyResultStatus,
            TextBlock proxyResultCount,
            Grid intermediaryResultRow,
            TextBlock intermediaryResultStatus,
            TextBlock intermediaryResultCount,
            Grid summaryResultRow,
            TextBlock summaryResultLabel,
            TextBlock summaryResultStatus,
            TextBlock summaryResultCount,
            Grid employeeRewardResultRow,
            TextBlock employeeRewardResultStatus,
            TextBlock employeeRewardResultCount,
            DockPanel finishedAtRow,
            TextBlock finishedAtText,
            Func<string, Brush> brushOf)
        {
            _completionCard = completionCard;
            _completionIconCircle = completionIconCircle;
            _completionIconText = completionIconText;
            _completionTitleText = completionTitleText;
            _completionDetailText = completionDetailText;
            _completionOutputLabel = completionOutputLabel;
            _completionOutputText = completionOutputText;
            _noProvinceResultHint = noProvinceResultHint;
            _stage1ResultRow = stage1ResultRow;
            _stage1ResultLabel = stage1ResultLabel;
            _stage1ResultStatus = stage1ResultStatus;
            _stage1ResultCount = stage1ResultCount;
            _proxyResultRow = proxyResultRow;
            _proxyResultStatus = proxyResultStatus;
            _proxyResultCount = proxyResultCount;
            _intermediaryResultRow = intermediaryResultRow;
            _intermediaryResultStatus = intermediaryResultStatus;
            _intermediaryResultCount = intermediaryResultCount;
            _summaryResultRow = summaryResultRow;
            _summaryResultLabel = summaryResultLabel;
            _summaryResultStatus = summaryResultStatus;
            _summaryResultCount = summaryResultCount;
            _employeeRewardResultRow = employeeRewardResultRow;
            _employeeRewardResultStatus = employeeRewardResultStatus;
            _employeeRewardResultCount = employeeRewardResultCount;
            _finishedAtRow = finishedAtRow;
            _finishedAtText = finishedAtText;
            _brushOf = brushOf;
        }

        public string LastOutputDirectory { get; private set; }

        public void UpdateResultVisibility(ProvinceCode? province)
        {
            var hasProvince = province.HasValue;
            var profile = hasProvince ? ProvinceUiProfile.For(province.Value) : null;
            var stageTwoVisibility = hasProvince && profile.SupportsStage2 ? Visibility.Visible : Visibility.Collapsed;
            var employeeRewardVisibility = hasProvince && profile.SupportsEmployeeReward ? Visibility.Visible : Visibility.Collapsed;

            _noProvinceResultHint.Visibility = hasProvince ? Visibility.Collapsed : Visibility.Visible;
            _stage1ResultRow.Visibility = hasProvince
                && (profile.SupportsStage1LedgerUpdate || profile.SupportsStage1CleanPower)
                    ? Visibility.Visible
                    : Visibility.Collapsed;
            _proxyResultRow.Visibility = stageTwoVisibility;
            _intermediaryResultRow.Visibility = stageTwoVisibility;
            _summaryResultRow.Visibility = stageTwoVisibility;
            _employeeRewardResultRow.Visibility = employeeRewardVisibility;
            _finishedAtRow.Visibility = hasProvince ? Visibility.Visible : Visibility.Collapsed;

            _stage1ResultLabel.Text = hasProvince ? profile.StageOneResultLabel : "阶段输出";
            _summaryResultLabel.Text = hasProvince && profile.StageTwo != null
                ? profile.StageTwo.ThirdResultLabel
                : "汇总表";
            _completionOutputLabel.Text = "输出文件夹";
        }

        public void Reset(ProvinceCode? province)
        {
            _stage1ResultStatus.Text = "等待";
            _stage1ResultCount.Text = "-";
            SetResultStatus(_proxyResultStatus, "等待");
            _proxyResultCount.Text = "-";
            SetResultStatus(_intermediaryResultStatus, "等待");
            _intermediaryResultCount.Text = "-";
            SetResultStatus(_summaryResultStatus, "等待");
            _summaryResultCount.Text = "-";
            _employeeRewardResultStatus.Text = "等待";
            _employeeRewardResultCount.Text = "-";
            _finishedAtText.Text = "-";
            LastOutputDirectory = null;
            ShowWaiting(province);
            _completionCard.Visibility = Visibility.Visible;
        }

        public void ShowWaiting(ProvinceCode? province)
        {
            var hasProvince = province.HasValue;
            ResetCompletionAppearance();
            _completionIconCircle.Background = hasProvince ? _brushOf("SuccessBrush") : _brushOf("AccentBrush");
            _completionIconText.Text = hasProvince ? "\uE73E" : "\uE946";
            _completionTitleText.Text = hasProvince ? "等待生成结果" : "等待选择省份";
            _completionDetailText.Text = hasProvince ? "运行完成后会在这里显示输出位置" : "选择结算省份后会显示对应输出项";
            _completionOutputText.Text = hasProvince ? "尚未生成" : "尚未选择省份";
            UpdateResultVisibility(province);
        }

        public void SetStage1Success(string countText)
        {
            _stage1ResultStatus.Text = "成功";
            _stage1ResultCount.Text = countText;
            SetFinishedNow();
        }

        public void SetStage2Success(string proxyCountText, string intermediaryCountText, string summaryCountText)
        {
            SetStage2Outcome(
                "成功",
                proxyCountText,
                "成功",
                intermediaryCountText,
                "成功",
                summaryCountText);
        }

        public void SetStage2Outcome(
            string proxyStatus,
            string proxyCountText,
            string intermediaryStatus,
            string intermediaryCountText,
            string summaryStatus,
            string summaryCountText)
        {
            SetResultStatus(_proxyResultStatus, proxyStatus);
            _proxyResultCount.Text = proxyCountText;
            SetResultStatus(_intermediaryResultStatus, intermediaryStatus);
            _intermediaryResultCount.Text = intermediaryCountText;
            SetResultStatus(_summaryResultStatus, summaryStatus);
            _summaryResultCount.Text = summaryCountText;
            SetFinishedNow();
        }

        public void SetStage2PreflightSuccess()
        {
            SetResultStatus(_proxyResultStatus, "预检");
            _proxyResultCount.Text = "未生成";
            SetResultStatus(_intermediaryResultStatus, "预检");
            _intermediaryResultCount.Text = "未生成";
            SetResultStatus(_summaryResultStatus, "预检");
            _summaryResultCount.Text = "未生成";
            SetFinishedNow();
        }

        public void SetEmployeeRewardSuccess(string personalCountText, string summaryCountText)
        {
            _employeeRewardResultStatus.Text = "成功";
            _employeeRewardResultCount.Text = personalCountText;
            _summaryResultStatus.Text = "成功";
            _summaryResultCount.Text = summaryCountText;
            SetFinishedNow();
        }

        public void ShowCompletion(string title, string detail, string outputDirectory)
        {
            LastOutputDirectory = outputDirectory;
            ResetCompletionAppearance();
            _completionIconCircle.Background = _brushOf("SuccessBrush");
            _completionIconText.Text = "\uE73E";
            _completionTitleText.Text = title;
            _completionDetailText.Text = detail;
            _completionOutputText.Text = outputDirectory;
            _completionCard.Visibility = Visibility.Visible;
        }

        public void ShowReviewCompletion(
            string title,
            string detail,
            string outputDirectory,
            bool critical)
        {
            LastOutputDirectory = outputDirectory;
            _completionCard.SetResourceReference(Border.BackgroundProperty, "StatusBusyBrush");
            _completionCard.SetResourceReference(
                Border.BorderBrushProperty,
                critical ? "ErrorBrush" : "WarningBrush");
            _completionIconCircle.SetResourceReference(
                Border.BackgroundProperty,
                critical ? "ErrorBrush" : "WarningBrush");
            _completionIconText.Text = critical ? "\uE783" : "\uE7BA";
            _completionTitleText.Text = title;
            _completionDetailText.Text = detail;
            _completionOutputText.Text = outputDirectory;
            _completionCard.Visibility = Visibility.Visible;
        }

        private void SetResultStatus(TextBlock statusText, string status)
        {
            var failed = string.Equals(status, "失败", StringComparison.Ordinal);
            var needsReview = failed || string.Equals(status, "需复核", StringComparison.Ordinal);
            statusText.Text = status;
            statusText.SetResourceReference(
                TextBlock.ForegroundProperty,
                failed ? "ErrorBrush" : needsReview ? "WarningBrush" : "SuccessBrush");
            var pill = statusText.Parent as Border;
            if (pill != null)
            {
                pill.SetResourceReference(
                    Border.BackgroundProperty,
                    needsReview ? "StatusBusyBrush" : "ResultPillBrush");
            }
        }

        private void ResetCompletionAppearance()
        {
            _completionCard.SetResourceReference(Border.BackgroundProperty, "CompletionBrush");
            _completionCard.SetResourceReference(Border.BorderBrushProperty, "CompletionBorderBrush");
        }

        private void SetFinishedNow()
        {
            _finishedAtText.Text = DateTime.Now.ToString("HH:mm:ss");
        }
    }
}
