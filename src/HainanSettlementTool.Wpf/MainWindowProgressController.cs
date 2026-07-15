using System;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Shapes;

namespace HainanSettlementTool.Wpf
{
    internal sealed class MainWindowProgressController
    {
        private readonly TextBlock _statusText;
        private readonly Shape _statusDot;
        private readonly Border _statusPill;
        private readonly TextBlock _progressTitle;
        private readonly TextBlock _progressDescriptionText;
        private readonly ProgressBar _progressBar;
        private readonly TextBlock _progressPercentText;
        private readonly TextBlock[] _stepTexts;
        private readonly TextBlock[] _stepStatuses;
        private readonly Func<string, Brush> _brushOf;
        private readonly string[] _defaultStepNames =
        {
            "检查输入文件",
            "读取台账数据",
            "生成结算文件",
            "写入结果报告",
            "保存结果文件"
        };
        private string[] _stepNames;

        public MainWindowProgressController(
            TextBlock statusText,
            Shape statusDot,
            Border statusPill,
            TextBlock progressTitle,
            TextBlock progressDescriptionText,
            ProgressBar progressBar,
            TextBlock progressPercentText,
            TextBlock[] stepTexts,
            TextBlock[] stepStatuses,
            Func<string, Brush> brushOf)
        {
            _statusText = statusText;
            _statusDot = statusDot;
            _statusPill = statusPill;
            _progressTitle = progressTitle;
            _progressDescriptionText = progressDescriptionText;
            _progressBar = progressBar;
            _progressPercentText = progressPercentText;
            _stepTexts = stepTexts;
            _stepStatuses = stepStatuses;
            _brushOf = brushOf;
            _stepNames = _defaultStepNames;
        }

        public void SetStatus(string text, string dotBrushKey, string backgroundBrushKey)
        {
            _statusText.Text = text;
            _statusDot.Fill = _brushOf(dotBrushKey);
            _statusPill.Background = _brushOf(backgroundBrushKey);
        }

        public void RefreshStatusBrushes()
        {
            if (_statusText.Text == "执行异常")
            {
                SetStatus(_statusText.Text, "ErrorBrush", "StatusBusyBrush");
            }
            else if (_statusText.Text == "待确认"
                || _statusText.Text == "运行中"
                || _statusText.Text == "需复核")
            {
                SetStatus(_statusText.Text, "WarningBrush", "StatusBusyBrush");
            }
            else
            {
                SetStatus(_statusText.Text, "SuccessBrush", "StatusReadyBrush");
            }
        }

        public void ResetProgress(string title, string description)
        {
            ResetProgress(title, description, null);
        }

        public void ResetProgress(string title, string description, string[] stepNames)
        {
            _stepNames = stepNames != null && stepNames.Length == _stepTexts.Length
                ? stepNames
                : _defaultStepNames;
            _progressTitle.Text = title;
            _progressDescriptionText.Text = description;
            SetProgress(0, description);
            for (var i = 0; i < _stepTexts.Length; i++)
            {
                SetStepWaiting(i);
            }
        }

        public void SetProgress(int value, string description)
        {
            _progressBar.Value = value;
            _progressPercentText.Text = value + "%";
            _progressDescriptionText.Text = description;
        }

        public void SetStepWaiting(int index)
        {
            _stepTexts[index].Text = "○  " + StepName(index);
            _stepTexts[index].Foreground = _brushOf("MutedBrush");
            _stepStatuses[index].Text = "等待中";
            _stepStatuses[index].Foreground = _brushOf("MutedBrush");
        }

        public void SetStepRunning(int index)
        {
            _stepTexts[index].Text = "●  " + StepName(index);
            _stepTexts[index].Foreground = _brushOf("AccentBrush");
            _stepStatuses[index].Text = "进行中";
            _stepStatuses[index].Foreground = _brushOf("AccentBrush");
        }

        public void SetStepNeedsConfirmation(int index)
        {
            _stepTexts[index].Text = "●  " + StepName(index);
            _stepTexts[index].Foreground = _brushOf("WarningBrush");
            _stepStatuses[index].Text = "待确认";
            _stepStatuses[index].Foreground = _brushOf("WarningBrush");
        }

        public void SetStepDone(int index)
        {
            _stepTexts[index].Text = "●  " + StepName(index);
            _stepTexts[index].Foreground = _brushOf("SuccessBrush");
            _stepStatuses[index].Text = "完成";
            _stepStatuses[index].Foreground = _brushOf("SuccessBrush");

            if (index + 1 < _stepTexts.Length)
            {
                SetStepRunning(index + 1);
                SetProgress(Math.Min(90, 25 + (index + 1) * 15), StepName(index + 1));
            }
        }

        public void SetStepFailed()
        {
            for (var i = 0; i < _stepStatuses.Length; i++)
            {
                if (_stepStatuses[i].Text == "进行中")
                {
                    _stepTexts[i].Text = "●  " + StepName(i);
                    _stepTexts[i].Foreground = _brushOf("ErrorBrush");
                    _stepStatuses[i].Text = "失败";
                    _stepStatuses[i].Foreground = _brushOf("ErrorBrush");
                    return;
                }
            }
        }

        private string StepName(int index)
        {
            return index >= 0 && index < _stepNames.Length ? _stepNames[index] : "处理文件";
        }
    }
}
