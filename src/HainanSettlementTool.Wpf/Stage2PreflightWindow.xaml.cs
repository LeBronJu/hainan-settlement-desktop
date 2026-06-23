using System.Linq;
using System.Text;
using System.Windows;
using HainanSettlementTool.Core.Models;

namespace HainanSettlementTool.Wpf
{
    public partial class Stage2PreflightWindow : Window
    {
        public Stage2PreflightWindow(Stage2PreflightReport report)
        {
            InitializeComponent();
            SummaryText.Text = "结算月份：2026年" + report.Month + "月；共发现 " + report.Issues.Count + " 条需要确认的变化。";
            DetailBox.Text = BuildDetailText(report);
        }

        private static string BuildDetailText(Stage2PreflightReport report)
        {
            var builder = new StringBuilder();
            var grouped = report.Issues
                .GroupBy(issue => issue.Category)
                .OrderBy(group => group.Key);

            foreach (var group in grouped)
            {
                builder.AppendLine("【" + group.Key + "】");
                var index = 1;
                foreach (var issue in group)
                {
                    builder.AppendLine(index + ". [" + issue.Severity + "] " + issue.Message);
                    if (!string.IsNullOrWhiteSpace(issue.Kind))
                    {
                        builder.AppendLine("   类型：" + issue.Kind);
                    }
                    if (!string.IsNullOrWhiteSpace(issue.Owner) || !string.IsNullOrWhiteSpace(issue.Entity))
                    {
                        builder.AppendLine("   负责人/主体：" + issue.Owner + " / " + issue.Entity);
                    }
                    if (!string.IsNullOrWhiteSpace(issue.Customer))
                    {
                        builder.AppendLine("   客户：" + issue.Customer);
                    }
                    if (issue.LedgerRow > 0)
                    {
                        builder.AppendLine("   台账行：" + issue.LedgerRow);
                    }
                    if (!string.IsNullOrWhiteSpace(issue.PreviousValue) || !string.IsNullOrWhiteSpace(issue.CurrentValue))
                    {
                        builder.AppendLine("   对比：" + issue.PreviousValue + "；" + issue.CurrentValue);
                    }
                    if (!string.IsNullOrWhiteSpace(issue.Suggestion))
                    {
                        builder.AppendLine("   建议：" + issue.Suggestion);
                    }
                    if (!string.IsNullOrWhiteSpace(issue.TemplateFile))
                    {
                        builder.AppendLine("   上月模板：" + issue.TemplateFile);
                    }
                    if (!string.IsNullOrWhiteSpace(issue.SheetName))
                    {
                        builder.AppendLine("   工作表：" + issue.SheetName);
                    }
                    builder.AppendLine();
                    index++;
                }
            }

            return builder.ToString();
        }

        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
    }
}
