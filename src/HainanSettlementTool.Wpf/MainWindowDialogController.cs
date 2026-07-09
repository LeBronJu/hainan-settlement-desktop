using System;
using System.Windows;

namespace HainanSettlementTool.Wpf
{
    internal sealed class MainWindowDialogController
    {
        private readonly Window _owner;

        public MainWindowDialogController(Window owner)
        {
            _owner = owner ?? throw new ArgumentNullException(nameof(owner));
        }

        public void ShowError(Exception ex)
        {
            ShowErrorMessage(ex.Message);
        }

        public void ShowErrorMessage(string message)
        {
            var dialog = new ModernDialogWindow("出错了", "需要先处理这个问题", message, "知道了", null, ModernDialogKind.Error)
            {
                Owner = _owner
            };
            dialog.ShowDialog();
        }

        public void ShowWarningMessage(string title, string heading, string message)
        {
            var dialog = new ModernDialogWindow(title, heading, message, "知道了", null, ModernDialogKind.Warning)
            {
                Owner = _owner
            };
            dialog.ShowDialog();
        }

        public bool ConfirmAction(string title, string heading, string message, string primaryButtonText)
        {
            var dialog = new ModernDialogWindow(title, heading, message, primaryButtonText, "取消", ModernDialogKind.Warning)
            {
                Owner = _owner
            };
            return dialog.ShowDialog() == true;
        }

        public bool ConfirmRun(string stageName, int month, string outputDirectory)
        {
            var dialog = new ConfirmRunWindow(stageName, month, outputDirectory)
            {
                Owner = _owner
            };
            return dialog.ShowDialog() == true;
        }
    }
}
