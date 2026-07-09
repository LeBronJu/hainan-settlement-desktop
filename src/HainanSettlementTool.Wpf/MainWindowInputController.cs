using System;
using System.IO;
using System.Windows.Controls;
using HainanSettlementTool.Core.Models;

namespace HainanSettlementTool.Wpf
{
    internal sealed class MainWindowInputController
    {
        private readonly ComboBox _monthCombo;
        private readonly ComboBox _rewardStartMonthCombo;
        private readonly ComboBox _rewardEndMonthCombo;
        private readonly ComboBox _provinceCombo;
        private readonly TextBox _outputDirBox;
        private readonly TextBox _baseLedgerBox;
        private readonly TextBox _powerBox;
        private readonly TextBox _rawDetailBox;
        private readonly TextBox _referenceLedgerBox;
        private readonly TextBox _completedLedgerBox;
        private readonly TextBox _proxyTemplateDirBox;
        private readonly TextBox _intermediaryTemplateDirBox;
        private readonly TextBox _refundTemplateDirBox;
        private readonly TextBox _summaryTemplateBox;
        private readonly TextBox _rewardLedgerBox;
        private readonly CheckBox _copyReferenceExistingCheckBox;
        private readonly CheckBox _allowMissingOwnerCheckBox;

        public MainWindowInputController(
            ComboBox monthCombo,
            ComboBox rewardStartMonthCombo,
            ComboBox rewardEndMonthCombo,
            ComboBox provinceCombo,
            TextBox outputDirBox,
            TextBox baseLedgerBox,
            TextBox powerBox,
            TextBox rawDetailBox,
            TextBox referenceLedgerBox,
            TextBox completedLedgerBox,
            TextBox proxyTemplateDirBox,
            TextBox intermediaryTemplateDirBox,
            TextBox refundTemplateDirBox,
            TextBox summaryTemplateBox,
            TextBox rewardLedgerBox,
            CheckBox copyReferenceExistingCheckBox,
            CheckBox allowMissingOwnerCheckBox)
        {
            _monthCombo = monthCombo;
            _rewardStartMonthCombo = rewardStartMonthCombo;
            _rewardEndMonthCombo = rewardEndMonthCombo;
            _provinceCombo = provinceCombo;
            _outputDirBox = outputDirBox;
            _baseLedgerBox = baseLedgerBox;
            _powerBox = powerBox;
            _rawDetailBox = rawDetailBox;
            _referenceLedgerBox = referenceLedgerBox;
            _completedLedgerBox = completedLedgerBox;
            _proxyTemplateDirBox = proxyTemplateDirBox;
            _intermediaryTemplateDirBox = intermediaryTemplateDirBox;
            _refundTemplateDirBox = refundTemplateDirBox;
            _summaryTemplateBox = summaryTemplateBox;
            _rewardLedgerBox = rewardLedgerBox;
            _copyReferenceExistingCheckBox = copyReferenceExistingCheckBox;
            _allowMissingOwnerCheckBox = allowMissingOwnerCheckBox;
        }

        public void LoadSavedInputs(UserInputSnapshot snapshot)
        {
            _outputDirBox.Text = snapshot.OutputDirectory ?? string.Empty;
            _baseLedgerBox.Text = snapshot.BaseLedgerPath ?? string.Empty;
            _powerBox.Text = snapshot.PowerPath ?? string.Empty;
            _rawDetailBox.Text = snapshot.RawDetailPath ?? string.Empty;
            _referenceLedgerBox.Text = snapshot.ReferenceLedgerPath ?? string.Empty;
            _completedLedgerBox.Text = snapshot.CompletedLedgerPath ?? string.Empty;
            _proxyTemplateDirBox.Text = snapshot.ProxyTemplateDirectory ?? string.Empty;
            _intermediaryTemplateDirBox.Text = snapshot.IntermediaryTemplateDirectory ?? string.Empty;
            _refundTemplateDirBox.Text = snapshot.RefundTemplateDirectory ?? string.Empty;
            _summaryTemplateBox.Text = snapshot.SummaryTemplatePath ?? string.Empty;
            _rewardLedgerBox.Text = snapshot.RewardLedgerPath ?? string.Empty;
        }

        public void SaveInputs(string themeMode)
        {
            UserInputStore.Save(new UserInputSnapshot
            {
                OutputDirectory = _outputDirBox.Text.Trim(),
                BaseLedgerPath = _baseLedgerBox.Text.Trim(),
                PowerPath = _powerBox.Text.Trim(),
                RawDetailPath = _rawDetailBox.Text.Trim(),
                ReferenceLedgerPath = _referenceLedgerBox.Text.Trim(),
                CompletedLedgerPath = _completedLedgerBox.Text.Trim(),
                ProxyTemplateDirectory = _proxyTemplateDirBox.Text.Trim(),
                IntermediaryTemplateDirectory = _intermediaryTemplateDirBox.Text.Trim(),
                RefundTemplateDirectory = _refundTemplateDirBox.Text.Trim(),
                SummaryTemplatePath = _summaryTemplateBox.Text.Trim(),
                RewardLedgerPath = _rewardLedgerBox.Text.Trim(),
                ProvinceCode = SelectedProvinceOrNull()?.ToString() ?? string.Empty,
                ThemeMode = themeMode
            });
        }

        public bool HasSavedInputs(UserInputSnapshot snapshot)
        {
            return !string.IsNullOrWhiteSpace(snapshot.OutputDirectory)
                || !string.IsNullOrWhiteSpace(snapshot.BaseLedgerPath)
                || !string.IsNullOrWhiteSpace(snapshot.PowerPath)
                || !string.IsNullOrWhiteSpace(snapshot.RawDetailPath)
                || !string.IsNullOrWhiteSpace(snapshot.ReferenceLedgerPath)
                || !string.IsNullOrWhiteSpace(snapshot.CompletedLedgerPath)
                || !string.IsNullOrWhiteSpace(snapshot.ProxyTemplateDirectory)
                || !string.IsNullOrWhiteSpace(snapshot.IntermediaryTemplateDirectory)
                || !string.IsNullOrWhiteSpace(snapshot.RefundTemplateDirectory)
                || !string.IsNullOrWhiteSpace(snapshot.SummaryTemplatePath)
                || !string.IsNullOrWhiteSpace(snapshot.RewardLedgerPath);
        }

        public HainanStage1Options CreateHainanStage1Options()
        {
            var powerPath = _powerBox.Text.Trim();
            var rawDetailPath = _rawDetailBox.Text.Trim();
            if (!string.IsNullOrWhiteSpace(rawDetailPath))
            {
                powerPath = ResolvePowerOutputPath(rawDetailPath);
                _powerBox.Text = powerPath;
            }

            return new HainanStage1Options
            {
                Month = SelectedMonth(),
                BaseLedgerPath = _baseLedgerBox.Text.Trim(),
                PowerPath = powerPath,
                RawDetailPath = rawDetailPath,
                ReferenceLedgerPath = _referenceLedgerBox.Text.Trim(),
                OutputDirectory = _outputDirBox.Text.Trim(),
                CopyReferenceExisting = _copyReferenceExistingCheckBox.IsChecked == true
            };
        }

        public string ResolvePowerOutputPath(string rawDetailPath)
        {
            if (string.IsNullOrWhiteSpace(rawDetailPath))
            {
                throw new InvalidOperationException("请选择原始零售侧明细。");
            }

            var outputDirectory = _outputDirBox.Text.Trim();
            if (string.IsNullOrWhiteSpace(outputDirectory))
            {
                throw new InvalidOperationException("请选择结果输出文件夹。");
            }

            return Path.Combine(outputDirectory, "零售侧用户电量数据处理表.xlsx");
        }

        public HainanPowerCleanInput PrepareHainanPowerCleanInput()
        {
            var rawDetailPath = _rawDetailBox.Text.Trim();
            var outputPath = ResolvePowerOutputPath(rawDetailPath);
            var outputDirectory = _outputDirBox.Text.Trim();
            _powerBox.Text = outputPath;
            return new HainanPowerCleanInput(rawDetailPath, outputPath, outputDirectory);
        }

        public HainanStage2Options CreateHainanStage2Options()
        {
            return new HainanStage2Options
            {
                Month = SelectedMonth(),
                LedgerPath = _completedLedgerBox.Text.Trim(),
                ProxyTemplateDirectory = _proxyTemplateDirBox.Text.Trim(),
                IntermediaryTemplateDirectory = _intermediaryTemplateDirBox.Text.Trim(),
                SummaryTemplatePath = _summaryTemplateBox.Text.Trim(),
                OutputDirectory = _outputDirBox.Text.Trim(),
                AllowMissingOwner = _allowMissingOwnerCheckBox.IsChecked == true
            };
        }

        public ChongqingStage2Options CreateChongqingStage2Options()
        {
            return new ChongqingStage2Options
            {
                Month = SelectedMonth(),
                LedgerPath = _completedLedgerBox.Text.Trim(),
                ProxyTemplateDirectory = _proxyTemplateDirBox.Text.Trim(),
                IntermediaryTemplateDirectory = _intermediaryTemplateDirBox.Text.Trim(),
                RefundTemplateDirectory = _refundTemplateDirBox.Text.Trim(),
                SummaryTemplatePath = _summaryTemplateBox.Text.Trim(),
                OutputDirectory = _outputDirBox.Text.Trim()
            };
        }

        public HainanEmployeePowerRewardOptions CreateHainanEmployeePowerRewardOptions()
        {
            var startMonth = SelectedRewardStartMonth();
            var endMonth = SelectedRewardEndMonth();
            if (startMonth > endMonth)
            {
                throw new InvalidOperationException("员工电量奖励开始月份不能晚于结束月份。");
            }

            return new HainanEmployeePowerRewardOptions
            {
                Year = 2026,
                StartMonth = startMonth,
                EndMonth = endMonth,
                LedgerPath = _rewardLedgerBox.Text.Trim(),
                OutputDirectory = _outputDirBox.Text.Trim()
            };
        }

        public ProvinceStage1CleanOptions CreateProvinceStage1CleanOptions()
        {
            return new ProvinceStage1CleanOptions
            {
                Province = SelectedProvince(),
                Month = SelectedMonthOrZero(),
                RawDetailPath = _rawDetailBox.Text.Trim(),
                OutputDirectory = _outputDirBox.Text.Trim()
            };
        }

        public ProvinceStage1LedgerUpdateOptions CreateProvinceStage1LedgerUpdateOptions()
        {
            return new ProvinceStage1LedgerUpdateOptions
            {
                Province = SelectedProvince(),
                Month = SelectedMonth(),
                LedgerPath = _baseLedgerBox.Text.Trim(),
                RawDetailPath = _rawDetailBox.Text.Trim(),
                OutputDirectory = _outputDirBox.Text.Trim()
            };
        }

        public int SelectedMonth()
        {
            if (_monthCombo.SelectedIndex < 0)
            {
                throw new InvalidOperationException("请选择结算月份。");
            }

            return _monthCombo.SelectedIndex + 2;
        }

        public int SelectedMonthOrZero()
        {
            return _monthCombo.SelectedIndex < 0 ? 0 : _monthCombo.SelectedIndex + 2;
        }

        public int SelectedRewardStartMonth()
        {
            if (_rewardStartMonthCombo.SelectedIndex < 0)
            {
                throw new InvalidOperationException("请选择员工电量奖励开始月份。");
            }

            return _rewardStartMonthCombo.SelectedIndex + 1;
        }

        public int SelectedRewardEndMonth()
        {
            if (_rewardEndMonthCombo.SelectedIndex < 0)
            {
                throw new InvalidOperationException("请选择员工电量奖励结束月份。");
            }

            return _rewardEndMonthCombo.SelectedIndex + 1;
        }

        public ProvinceCode SelectedProvince()
        {
            var province = SelectedProvinceOrNull();
            if (!province.HasValue)
            {
                throw new InvalidOperationException("请选择结算省份。");
            }

            return province.Value;
        }

        public ProvinceCode? SelectedProvinceOrNull()
        {
            return SelectedProfileOrNull()?.Province;
        }

        public ProvinceUiProfile SelectedProfileOrNull()
        {
            return _provinceCombo.SelectedItem as ProvinceUiProfile;
        }

        public void ClearStage1()
        {
            _baseLedgerBox.Clear();
            _powerBox.Clear();
            _rawDetailBox.Clear();
            _referenceLedgerBox.Clear();
            _copyReferenceExistingCheckBox.IsChecked = false;
        }

        public void ClearStage2()
        {
            _completedLedgerBox.Clear();
            _proxyTemplateDirBox.Clear();
            _intermediaryTemplateDirBox.Clear();
            _refundTemplateDirBox.Clear();
            _summaryTemplateBox.Clear();
            _allowMissingOwnerCheckBox.IsChecked = false;
        }

        public void ClearEmployeeReward()
        {
            _rewardLedgerBox.Clear();
            _rewardStartMonthCombo.SelectedIndex = 0;
            _rewardEndMonthCombo.SelectedIndex = -1;
        }
    }

    internal sealed class HainanPowerCleanInput
    {
        public HainanPowerCleanInput(string rawDetailPath, string outputPath, string outputDirectory)
        {
            RawDetailPath = rawDetailPath;
            OutputPath = outputPath;
            OutputDirectory = outputDirectory;
        }

        public string RawDetailPath { get; }

        public string OutputPath { get; }

        public string OutputDirectory { get; }
    }
}
