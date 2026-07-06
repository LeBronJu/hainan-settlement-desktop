using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace HainanSettlementTool.Wpf
{
    public enum ModernDialogKind
    {
        Info,
        Warning,
        Error
    }

    public partial class ModernDialogWindow : Window
    {
        public ModernDialogWindow(
            string title,
            string heading,
            string message,
            string primaryButtonText,
            string secondaryButtonText,
            ModernDialogKind kind)
        {
            InitializeComponent();
            Title = title;
            DialogTitleText.Text = title;
            HeadingText.Text = heading;
            MessageText.Text = message;
            PrimaryButton.Content = primaryButtonText;

            if (string.IsNullOrWhiteSpace(secondaryButtonText))
            {
                SecondaryButton.Visibility = Visibility.Collapsed;
                PrimaryButton.Width = 104;
            }
            else
            {
                SecondaryButton.Content = secondaryButtonText;
            }

            ApplyKind(kind);
        }

        private void ApplyKind(ModernDialogKind kind)
        {
            switch (kind)
            {
                case ModernDialogKind.Error:
                    IconText.Text = "\uE783";
                    IconText.SetResourceReference(TextBlock.ForegroundProperty, "ErrorBrush");
                    break;
                case ModernDialogKind.Warning:
                    IconText.Text = "\uE7BA";
                    IconText.SetResourceReference(TextBlock.ForegroundProperty, "WarningBrush");
                    break;
                default:
                    IconText.Text = "\uE946";
                    IconText.SetResourceReference(TextBlock.ForegroundProperty, "AccentBrush");
                    break;
            }

            IconCircle.SetResourceReference(Border.BackgroundProperty, "PanelSoftBrush");
        }

        private void Primary_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
        }

        private void Secondary_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }

        private void Chrome_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ButtonState == MouseButtonState.Pressed)
            {
                DragMove();
            }
        }
    }
}
