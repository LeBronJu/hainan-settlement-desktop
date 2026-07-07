using System;
using System.Collections.Generic;
using System.Linq;
using HainanSettlementTool.Core.Models;

namespace HainanSettlementTool.Wpf
{
    internal sealed class ProvinceUiProfile
    {
        private static readonly ProvinceUiProfile HainanProfile = new ProvinceUiProfile(
            ProvinceCode.Hainan,
            "海南",
            "代理费结算",
            "阶段一：写入电量到台账",
            "整理电量并写入基础台账，输出检查报告",
            "基础台账（必填）",
            "电量处理表",
            "原始零售侧明细",
            "参考台账（可选）",
            "开始 执行阶段一",
            "只清洗电量",
            "阶段一和阶段二生成的所有文件都会保存到这个文件夹中",
            "阶段一报告",
            "选择基础台账",
            "选择原始零售侧明细",
            supportsStage1LedgerUpdate: true,
            supportsStage1CleanPower: true,
            supportsStage2: true,
            supportsEmployeeReward: true,
            showsExistingPowerInput: true,
            showsReferenceLedgerInput: true);

        private static readonly ProvinceUiProfile ChongqingProfile = new ProvinceUiProfile(
            ProvinceCode.Chongqing,
            "重庆",
            "阶段一：台账更新",
            "阶段一：台账更新",
            "清洗交易中心电量确认结算单，并写入重庆台账副本",
            "重庆售电结算台账（必填）",
            "电量处理表",
            "交易中心电量确认结算单（必填）",
            "参考台账（可选）",
            "清洗并更新台账",
            "只清洗电量数据",
            "重庆当前开放阶段一的只清洗电量数据，输出仍保存到这个文件夹中",
            "阶段一台账更新",
            "选择重庆售电结算台账",
            "选择重庆交易中心电量确认结算单",
            supportsStage1LedgerUpdate: true,
            supportsStage1CleanPower: true,
            supportsStage2: false,
            supportsEmployeeReward: false,
            showsExistingPowerInput: false,
            showsReferenceLedgerInput: false);

        private ProvinceUiProfile(
            ProvinceCode province,
            string displayName,
            string mainSettlementTabHeader,
            string stageOneTitle,
            string stageOneCaption,
            string baseLedgerLabel,
            string existingPowerLabel,
            string rawDetailLabel,
            string referenceLedgerLabel,
            string runStageOneButtonText,
            string cleanPowerButtonText,
            string sharedSettingsCaption,
            string stageOneResultLabel,
            string baseLedgerDialogTitle,
            string rawDetailDialogTitle,
            bool supportsStage1LedgerUpdate,
            bool supportsStage1CleanPower,
            bool supportsStage2,
            bool supportsEmployeeReward,
            bool showsExistingPowerInput,
            bool showsReferenceLedgerInput)
        {
            Province = province;
            DisplayName = displayName;
            MainSettlementTabHeader = mainSettlementTabHeader;
            StageOneTitle = stageOneTitle;
            StageOneCaption = stageOneCaption;
            BaseLedgerLabel = baseLedgerLabel;
            ExistingPowerLabel = existingPowerLabel;
            RawDetailLabel = rawDetailLabel;
            ReferenceLedgerLabel = referenceLedgerLabel;
            RunStageOneButtonText = runStageOneButtonText;
            CleanPowerButtonText = cleanPowerButtonText;
            SharedSettingsCaption = sharedSettingsCaption;
            StageOneResultLabel = stageOneResultLabel;
            BaseLedgerDialogTitle = baseLedgerDialogTitle;
            RawDetailDialogTitle = rawDetailDialogTitle;
            SupportsStage1LedgerUpdate = supportsStage1LedgerUpdate;
            SupportsStage1CleanPower = supportsStage1CleanPower;
            SupportsStage2 = supportsStage2;
            SupportsEmployeeReward = supportsEmployeeReward;
            ShowsExistingPowerInput = showsExistingPowerInput;
            ShowsReferenceLedgerInput = showsReferenceLedgerInput;
        }

        public ProvinceCode Province { get; }

        public string DisplayName { get; }

        public string MainSettlementTabHeader { get; }

        public string StageOneTitle { get; }

        public string StageOneCaption { get; }

        public string BaseLedgerLabel { get; }

        public string ExistingPowerLabel { get; }

        public string RawDetailLabel { get; }

        public string ReferenceLedgerLabel { get; }

        public string RunStageOneButtonText { get; }

        public string CleanPowerButtonText { get; }

        public string SharedSettingsCaption { get; }

        public string StageOneResultLabel { get; }

        public string BaseLedgerDialogTitle { get; }

        public string RawDetailDialogTitle { get; }

        public bool SupportsStage1LedgerUpdate { get; }

        public bool SupportsStage1CleanPower { get; }

        public bool SupportsStage2 { get; }

        public bool SupportsEmployeeReward { get; }

        public bool ShowsExistingPowerInput { get; }

        public bool ShowsReferenceLedgerInput { get; }

        public static IReadOnlyList<ProvinceUiProfile> Supported { get; } =
            new[] { HainanProfile, ChongqingProfile };

        public static ProvinceUiProfile For(ProvinceCode province)
        {
            var profile = Supported.FirstOrDefault(item => item.Province == province);
            if (profile == null)
            {
                throw new ArgumentOutOfRangeException(nameof(province), province, "不支持的结算省份。");
            }

            return profile;
        }
    }
}
