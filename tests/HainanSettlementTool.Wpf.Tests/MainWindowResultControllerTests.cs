using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using HainanSettlementTool.Core.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace HainanSettlementTool.Wpf.Tests
{
    [TestClass]
    public sealed class MainWindowResultControllerTests
    {
        [STATestMethod]
        public void CompletionShowsReadableReportActionOnlyWhenReportPathExistsForRun()
        {
            var reportButton = new Button { Visibility = Visibility.Collapsed };
            var controller = CreateController(reportButton);

            controller.ShowCompletion("完成", "已生成", @"C:\output", @"C:\output\report.html");

            Assert.AreEqual(@"C:\output\report.html", controller.LastReadableReportPath);
            Assert.AreEqual(Visibility.Visible, reportButton.Visibility);

            controller.ShowCompletion("完成", "已生成", @"C:\output");

            Assert.IsNull(controller.LastReadableReportPath);
            Assert.AreEqual(Visibility.Collapsed, reportButton.Visibility);

            controller.ShowReviewCompletion(
                "需复核",
                "请检查",
                @"C:\output",
                false,
                @"C:\output\review.html");

            Assert.AreEqual(@"C:\output\review.html", controller.LastReadableReportPath);
            Assert.AreEqual(Visibility.Visible, reportButton.Visibility);

            controller.Reset(ProvinceCode.Hainan);

            Assert.IsNull(controller.LastReadableReportPath);
            Assert.AreEqual(Visibility.Collapsed, reportButton.Visibility);
        }

        private static MainWindowResultController CreateController(Button reportButton)
        {
            return new MainWindowResultController(
                new Border(),
                new Border(),
                new TextBlock(),
                new TextBlock(),
                new TextBlock(),
                new TextBlock(),
                new TextBlock(),
                reportButton,
                new TextBlock(),
                new Grid(),
                new TextBlock(),
                new TextBlock(),
                new TextBlock(),
                new Grid(),
                new TextBlock(),
                new TextBlock(),
                new Grid(),
                new TextBlock(),
                new TextBlock(),
                new Grid(),
                new TextBlock(),
                new TextBlock(),
                new TextBlock(),
                new Grid(),
                new TextBlock(),
                new TextBlock(),
                new DockPanel(),
                new TextBlock(),
                _ => Brushes.Green);
        }
    }
}
