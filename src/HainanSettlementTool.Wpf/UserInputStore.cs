using System;
using System.IO;
using System.Xml.Serialization;

namespace HainanSettlementTool.Wpf
{
    public sealed class UserInputSnapshot
    {
        public string OutputDirectory { get; set; }
        public string BaseLedgerPath { get; set; }
        public string PowerPath { get; set; }
        public string RawDetailPath { get; set; }
        public string ReferenceLedgerPath { get; set; }
        public string CompletedLedgerPath { get; set; }
        public string ProxyTemplateDirectory { get; set; }
        public string IntermediaryTemplateDirectory { get; set; }
        public string RefundTemplateDirectory { get; set; }
        public string SummaryTemplatePath { get; set; }
        public string RewardLedgerPath { get; set; }
        public string ProvinceCode { get; set; }
        public string ThemeMode { get; set; }
    }

    internal static class UserInputStore
    {
        private const string CurrentSettingsFolderName = "SettlementAutomationTool";
        private const string LegacySettingsFolderName = "HainanSettlementTool";
        private const string SettingsFileName = "modern-ui-inputs.xml";

        private static readonly XmlSerializer Serializer = new XmlSerializer(typeof(UserInputSnapshot));

        public static UserInputSnapshot Load()
        {
            var path = ReadSettingsPath();
            if (!File.Exists(path))
            {
                return new UserInputSnapshot();
            }

            try
            {
                using (var stream = File.OpenRead(path))
                {
                    return (UserInputSnapshot)Serializer.Deserialize(stream);
                }
            }
            catch
            {
                return new UserInputSnapshot();
            }
        }

        public static void Save(UserInputSnapshot snapshot)
        {
            var path = SettingsPath();
            Directory.CreateDirectory(Path.GetDirectoryName(path));

            using (var stream = File.Create(path))
            {
                Serializer.Serialize(stream, snapshot);
            }
        }

        private static string SettingsPath()
        {
            var root = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            return Path.Combine(root, CurrentSettingsFolderName, SettingsFileName);
        }

        private static string ReadSettingsPath()
        {
            var path = SettingsPath();
            if (File.Exists(path))
            {
                return path;
            }

            var root = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            var legacyPath = Path.Combine(root, LegacySettingsFolderName, SettingsFileName);
            return File.Exists(legacyPath) ? legacyPath : path;
        }
    }
}
