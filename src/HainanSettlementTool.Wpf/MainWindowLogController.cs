using System;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;

namespace HainanSettlementTool.Wpf
{
    internal sealed class MainWindowLogController
    {
        private readonly Window _owner;
        private readonly TextBox _logBox;

        public MainWindowLogController(Window owner, TextBox logBox)
        {
            _owner = owner ?? throw new ArgumentNullException(nameof(owner));
            _logBox = logBox ?? throw new ArgumentNullException(nameof(logBox));
        }

        public void Add(string message, string level)
        {
            _logBox.AppendText("[" + DateTime.Now.ToString("HH:mm:ss") + "] [" + level + "] " + message + Environment.NewLine);
            _logBox.ScrollToEnd();
        }

        public void Clear()
        {
            _logBox.Clear();
        }

        public void Save()
        {
            var dialog = new SaveFileDialog
            {
                Title = "保存运行日志",
                Filter = "文本文件|*.txt|所有文件|*.*",
                FileName = "售电结算运行日志.txt"
            };

            if (dialog.ShowDialog(_owner) == true)
            {
                File.WriteAllText(dialog.FileName, _logBox.Text, Encoding.UTF8);
            }
        }
    }
}
