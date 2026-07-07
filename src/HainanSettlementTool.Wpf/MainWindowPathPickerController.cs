using System;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using Ookii.Dialogs.Wpf;

namespace HainanSettlementTool.Wpf
{
    internal sealed class MainWindowPathPickerController
    {
        private readonly Window _owner;

        public MainWindowPathPickerController(Window owner)
        {
            _owner = owner ?? throw new ArgumentNullException(nameof(owner));
        }

        public bool BrowseExcel(TextBox target, string title)
        {
            return BrowseFile(target, title, "Excel 文件|*.xlsx|所有文件|*.*");
        }

        public bool BrowseFile(TextBox target, string title, string filter)
        {
            var dialog = new OpenFileDialog
            {
                Title = title,
                Filter = filter,
                CheckFileExists = true
            };

            if (dialog.ShowDialog(_owner) != true)
            {
                return false;
            }

            target.Text = dialog.FileName;
            return true;
        }

        public bool BrowseFolder(TextBox target, string title)
        {
            var dialog = new VistaFolderBrowserDialog
            {
                Description = title,
                UseDescriptionForTitle = true,
                ShowNewFolderButton = true
            };

            if (Directory.Exists(target.Text))
            {
                dialog.SelectedPath = target.Text;
            }

            if (dialog.ShowDialog(_owner) != true)
            {
                return false;
            }

            target.Text = dialog.SelectedPath;
            return true;
        }
    }
}
