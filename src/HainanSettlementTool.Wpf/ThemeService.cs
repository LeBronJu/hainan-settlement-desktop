using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Media;
using Microsoft.Win32;

namespace HainanSettlementTool.Wpf
{
    internal static class ThemeService
    {
        public const string SystemMode = "System";
        public const string LightMode = "Light";
        public const string DarkMode = "Dark";

        private static readonly Dictionary<string, string> LightPalette = new Dictionary<string, string>
        {
            { "PageBrush", "#F4F7F8" },
            { "PanelBrush", "#FFFFFF" },
            { "PanelSoftBrush", "#F8FBFB" },
            { "TextBrush", "#17202A" },
            { "MutedBrush", "#65717D" },
            { "BorderBrushSoft", "#DDE5E8" },
            { "AccentBrush", "#008A8D" },
            { "AccentDarkBrush", "#006F72" },
            { "OnAccentBrush", "#FFFFFF" },
            { "SuccessBrush", "#168A4A" },
            { "WarningBrush", "#DB8B00" },
            { "ErrorBrush", "#C0392B" },
            { "FieldTextBrush", "#34404A" },
            { "WindowBorderBrush", "#AEBBC4" },
            { "TitleBarBrush", "#FAFCFD" },
            { "InputBrush", "#FFFFFF" },
            { "InputHoverBrush", "#FFFFFF" },
            { "ButtonSecondaryBrush", "#FBFCFD" },
            { "ButtonSecondaryHoverBrush", "#F1F5F6" },
            { "ComboItemHoverBrush", "#EEF6F7" },
            { "ComboItemSelectedBrush", "#DFF1F2" },
            { "ComboBorderHoverBrush", "#B9C8CE" },
            { "ComboPopupBorderBrush", "#C9D6DC" },
            { "TabStripBrush", "#E9EFF2" },
            { "TabHoverBrush", "#F7FAFB" },
            { "TabSelectedBrush", "#FFFFFF" },
            { "TabSelectedBorderBrush", "#CFDCE1" },
            { "StageOneBorderBrush", "#BBDDE1" },
            { "StageOneHeaderBrush", "#F1FBFC" },
            { "StageOneHeaderBorderBrush", "#D5ECEE" },
            { "StageTwoBorderBrush", "#BFDCC9" },
            { "StageTwoHeaderBrush", "#F2FBF5" },
            { "StageTwoHeaderBorderBrush", "#DDEFE2" },
            { "RewardBorderBrush", "#C9D8EA" },
            { "RewardHeaderBrush", "#F4F8FE" },
            { "RewardHeaderBorderBrush", "#DCE7F5" },
            { "InfoPanelBrush", "#FBFCFD" },
            { "LogBrush", "#17202A" },
            { "LogTextBrush", "#D8E7EC" },
            { "LogActionBrush", "#D7E4EA" },
            { "ProgressTrackBrush", "#E4EAED" },
            { "CompletionBrush", "#F7FCF8" },
            { "CompletionBorderBrush", "#A9D9BB" },
            { "CompletionInnerBrush", "#FFFFFF" },
            { "CompletionInnerBorderBrush", "#DDEFE2" },
            { "ResultPillBrush", "#ECF8F0" },
            { "StatusReadyBrush", "#EAF7F1" },
            { "StatusBusyBrush", "#FFF5E0" }
        };

        private static readonly Dictionary<string, string> DarkPalette = new Dictionary<string, string>
        {
            { "PageBrush", "#0F151A" },
            { "PanelBrush", "#151D23" },
            { "PanelSoftBrush", "#19242B" },
            { "TextBrush", "#E8EEF2" },
            { "MutedBrush", "#9AA9B4" },
            { "BorderBrushSoft", "#2D3B44" },
            { "AccentBrush", "#19A7A9" },
            { "AccentDarkBrush", "#118688" },
            { "OnAccentBrush", "#FFFFFF" },
            { "SuccessBrush", "#36B66A" },
            { "WarningBrush", "#F2A93B" },
            { "ErrorBrush", "#E05A4F" },
            { "FieldTextBrush", "#C9D4DC" },
            { "WindowBorderBrush", "#33434D" },
            { "TitleBarBrush", "#121A20" },
            { "InputBrush", "#101820" },
            { "InputHoverBrush", "#14202A" },
            { "ButtonSecondaryBrush", "#16212A" },
            { "ButtonSecondaryHoverBrush", "#1C2A34" },
            { "ComboItemHoverBrush", "#1D3038" },
            { "ComboItemSelectedBrush", "#143E44" },
            { "ComboBorderHoverBrush", "#49616D" },
            { "ComboPopupBorderBrush", "#394C56" },
            { "TabStripBrush", "#131C23" },
            { "TabHoverBrush", "#1B2830" },
            { "TabSelectedBrush", "#22313A" },
            { "TabSelectedBorderBrush", "#3A505C" },
            { "StageOneBorderBrush", "#236065" },
            { "StageOneHeaderBrush", "#13292D" },
            { "StageOneHeaderBorderBrush", "#254A50" },
            { "StageTwoBorderBrush", "#2D6041" },
            { "StageTwoHeaderBrush", "#14271C" },
            { "StageTwoHeaderBorderBrush", "#2C4A35" },
            { "RewardBorderBrush", "#405875" },
            { "RewardHeaderBrush", "#172235" },
            { "RewardHeaderBorderBrush", "#31445B" },
            { "InfoPanelBrush", "#111B23" },
            { "LogBrush", "#0A0F14" },
            { "LogTextBrush", "#D8E7EC" },
            { "LogActionBrush", "#BFD2DA" },
            { "ProgressTrackBrush", "#26353F" },
            { "CompletionBrush", "#13241A" },
            { "CompletionBorderBrush", "#2F6844" },
            { "CompletionInnerBrush", "#101A14" },
            { "CompletionInnerBorderBrush", "#2B4936" },
            { "ResultPillBrush", "#14351F" },
            { "StatusReadyBrush", "#153823" },
            { "StatusBusyBrush", "#3B2A13" }
        };

        public static string NormalizeMode(string mode)
        {
            if (string.Equals(mode, LightMode, StringComparison.OrdinalIgnoreCase))
            {
                return LightMode;
            }

            if (string.Equals(mode, DarkMode, StringComparison.OrdinalIgnoreCase))
            {
                return DarkMode;
            }

            return SystemMode;
        }

        public static void Apply(string mode)
        {
            var normalized = NormalizeMode(mode);
            var dark = normalized == DarkMode || (normalized == SystemMode && IsSystemDark());
            ApplyPalette(dark ? DarkPalette : LightPalette);
        }

        public static bool IsSystemDark()
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"))
                {
                    var value = key?.GetValue("AppsUseLightTheme");
                    if (value is int)
                    {
                        return (int)value == 0;
                    }
                }
            }
            catch
            {
            }

            return false;
        }

        private static void ApplyPalette(IDictionary<string, string> palette)
        {
            foreach (var item in palette)
            {
                Application.Current.Resources[item.Key] = BrushFromHex(item.Value);
            }
        }

        private static SolidColorBrush BrushFromHex(string hex)
        {
            return new SolidColorBrush((Color)ColorConverter.ConvertFromString(hex));
        }
    }
}
