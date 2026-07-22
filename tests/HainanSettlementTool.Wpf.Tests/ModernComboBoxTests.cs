using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace HainanSettlementTool.Wpf.Tests
{
    [TestClass]
    public class ModernComboBoxTests
    {
        [STATestMethod]
        public void SelectedTemplateOptionDisplaysFriendlyText()
        {
            EnsureApplicationResources();
            var option = new Stage2PreflightTemplateOptionViewModel
            {
                DisplayText = "模板主体 / 测试负责人 / 模板.xlsx",
                Path = @"C:\SyntheticTemplates\template.xlsx"
            };
            var comboBox = new ComboBox
            {
                ItemsSource = new[] { option },
                DisplayMemberPath = nameof(Stage2PreflightTemplateOptionViewModel.DisplayText),
                SelectedValuePath = nameof(Stage2PreflightTemplateOptionViewModel.Path),
                Style = (Style)Application.Current.FindResource("ModernComboBox"),
                Width = 620
            };

            comboBox.SelectedValue = option.Path;
            comboBox.Measure(new Size(620, 40));
            comboBox.Arrange(new Rect(0, 0, 620, 40));
            comboBox.ApplyTemplate();
            comboBox.UpdateLayout();

            var contentSite = (ContentPresenter)comboBox.Template.FindName("ContentSite", comboBox);
            Assert.IsNotNull(contentSite);
            contentSite.ApplyTemplate();
            contentSite.Measure(new Size(590, 32));
            contentSite.Arrange(new Rect(0, 0, 590, 32));
            contentSite.UpdateLayout();

            var renderedText = Descendants<TextBlock>(contentSite)
                .Select(textBlock => textBlock.Text)
                .FirstOrDefault(text => !string.IsNullOrWhiteSpace(text));

            Assert.AreEqual(option.DisplayText, renderedText);
        }

        private static void EnsureApplicationResources()
        {
            if (Application.Current != null)
            {
                return;
            }

            var application = new App();
            application.InitializeComponent();
        }

        private static IEnumerable<T> Descendants<T>(DependencyObject root)
            where T : DependencyObject
        {
            for (var index = 0; index < VisualTreeHelper.GetChildrenCount(root); index++)
            {
                var child = VisualTreeHelper.GetChild(root, index);
                var match = child as T;
                if (match != null)
                {
                    yield return match;
                }

                foreach (var descendant in Descendants<T>(child))
                {
                    yield return descendant;
                }
            }
        }
    }
}
