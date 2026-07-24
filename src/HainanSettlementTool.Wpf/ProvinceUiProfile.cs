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
            supportsEmployeeReward: true,
            showEmployeeRewardPlaceholder: false,
            showsExistingPowerInput: true,
            showsReferenceLedgerInput: true,
            stageTwo: new ProvinceStage2UiProfile(
                "阶段二：生成分表和汇总表",
                "生成代理/居间分表和汇总表，输出结算结果",
                "开始 执行阶段二",
                "人工整理后的台账（必填）",
                "上月代理分表文件夹",
                "上月居间分表文件夹",
                "退补分表文件夹",
                "上月/修正版汇总表（必填）",
                "汇总表",
                showsCompletedLedger: true,
                showsProxyDirectory: true,
                showsIntermediaryDirectory: true,
                showsRefundDirectory: false,
                showsSummaryTemplate: true,
                showsAllowMissingOwner: false));

        private static readonly ProvinceUiProfile ChongqingProfile = new ProvinceUiProfile(
            ProvinceCode.Chongqing,
            "重庆",
            "代理费结算",
            "阶段一：台账更新",
            "清洗交易中心电量确认结算单，并写入重庆台账副本",
            "重庆售电结算台账（必填）",
            "电量处理表",
            "交易中心电量确认结算单（必填）",
            "参考台账（可选）",
            "清洗并更新台账",
            "只清洗电量数据",
            "重庆当前开放阶段一生成和阶段二结算生成；阶段二会先预检再写出输出副本",
            "阶段一台账更新",
            "选择重庆售电结算台账",
            "选择重庆交易中心电量确认结算单",
            supportsStage1LedgerUpdate: true,
            supportsStage1CleanPower: true,
            supportsEmployeeReward: false,
            showEmployeeRewardPlaceholder: true,
            showsExistingPowerInput: false,
            showsReferenceLedgerInput: false,
            stageTwo: new ProvinceStage2UiProfile(
                "阶段二：重庆结算生成",
                "生成代理/居间/退补分表和汇总表，生成前先确认预检项目",
                "开始 重庆阶段二",
                "人工整理后的台账（必填）",
                "上月代理分表文件夹",
                "上月居间分表文件夹",
                "退补分表文件夹",
                "上月/修正版汇总表（必填）",
                "汇总表",
                showsCompletedLedger: true,
                showsProxyDirectory: true,
                showsIntermediaryDirectory: true,
                showsRefundDirectory: true,
                showsSummaryTemplate: true,
                showsAllowMissingOwner: false));

        private static readonly ProvinceUiProfile GuangdongProfile = new ProvinceUiProfile(
            ProvinceCode.Guangdong,
            "广东",
            "结算处理",
            "广东阶段一：整理本月电量",
            "整理交易中心明细，生成八列电量表并写入本月台账",
            "广东售电结算台账（必填）",
            "电量处理表",
            "交易中心零售结算明细（必填）",
            "参考台账",
            "生成本月台账",
            "整理本月电量",
            "生成的电量表、台账和检查报告都会保存到这个文件夹",
            "本月电量处理",
            "选择广东售电结算台账",
            "选择广东交易中心零售结算明细",
            supportsStage1LedgerUpdate: true,
            supportsStage1CleanPower: true,
            supportsEmployeeReward: false,
            showEmployeeRewardPlaceholder: false,
            showsExistingPowerInput: false,
            showsReferenceLedgerInput: false,
            stageTwo: new ProvinceStage2UiProfile(
                "广东分表月份初始化",
                "复制标准上月 sheet，或整理已有目标月 sheet，并清空 C-F 电量",
                "开始 初始化分表",
                "台账",
                "代理分表文件夹（可选）",
                "居间分表文件夹（可选）",
                "退补分表文件夹（可选）",
                "汇总表",
                "退补分表",
                showsCompletedLedger: false,
                showsProxyDirectory: true,
                showsIntermediaryDirectory: true,
                showsRefundDirectory: true,
                showsSummaryTemplate: false,
                showsAllowMissingOwner: false));

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
            bool supportsEmployeeReward,
            bool showEmployeeRewardPlaceholder,
            bool showsExistingPowerInput,
            bool showsReferenceLedgerInput,
            ProvinceStage2UiProfile stageTwo)
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
            SupportsEmployeeReward = supportsEmployeeReward;
            ShowEmployeeRewardPlaceholder = showEmployeeRewardPlaceholder;
            ShowsExistingPowerInput = showsExistingPowerInput;
            ShowsReferenceLedgerInput = showsReferenceLedgerInput;
            StageTwo = stageTwo;
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

        public bool SupportsStage2 => StageTwo != null;

        public bool SupportsEmployeeReward { get; }

        public bool ShowEmployeeRewardPlaceholder { get; }

        public bool ShowsExistingPowerInput { get; }

        public bool ShowsReferenceLedgerInput { get; }

        public ProvinceStage2UiProfile StageTwo { get; }

        public static IReadOnlyList<ProvinceUiProfile> Supported { get; } =
            new[] { HainanProfile, ChongqingProfile, GuangdongProfile };

        public static ProvinceUiProfile For(ProvinceCode province)
        {
            var profile = Supported.FirstOrDefault(item => item.Province == province);
            if (profile == null)
            {
                throw new ArgumentOutOfRangeException(nameof(province), province, "不支持的结算省份。");
            }

            return profile;
        }

        public override string ToString()
        {
            return DisplayName;
        }
    }
}
