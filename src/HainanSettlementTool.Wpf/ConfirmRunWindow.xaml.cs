using System.Windows;

namespace HainanSettlementTool.Wpf
{
    public partial class ConfirmRunWindow : Window
    {
        public ConfirmRunWindow(string stageName, int month, string outputDirectory)
        {
            InitializeComponent();
            TitleText.Text = "即将执行 " + stageName;
            MonthText.Text = "结算月份：" + "2026年" + month + "月";
            OutputText.Text = outputDirectory;
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
