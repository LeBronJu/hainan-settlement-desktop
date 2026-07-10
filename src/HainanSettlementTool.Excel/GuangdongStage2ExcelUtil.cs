using System;
using System.Globalization;
using System.IO;
using ClosedXML.Excel;
using HainanSettlementTool.Core.Models;

namespace HainanSettlementTool.Excel
{
    internal static class GuangdongStage2ExcelUtil
    {
        public static XLWorkbook OpenWorkbookShared(string path)
        {
            using (var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                return new XLWorkbook(stream);
            }
        }

        public static void CopyWorkbookShared(string source, string target)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(target));
            using (var input = File.Open(source, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var output = File.Create(target))
            {
                input.CopyTo(output);
            }
        }

        public static void SaveWorkbook(XLWorkbook workbook)
        {
            workbook.CalculateMode = XLCalculateMode.Auto;
            workbook.Save(new SaveOptions { EvaluateFormulasBeforeSaving = true });
        }

        public static string RelativePath(string root, string path)
        {
            var rootUri = new Uri(AppendDirectorySeparatorChar(Path.GetFullPath(root)));
            var pathUri = new Uri(Path.GetFullPath(path));
            return Uri.UnescapeDataString(rootUri.MakeRelativeUri(pathUri).ToString())
                .Replace('/', Path.DirectorySeparatorChar);
        }

        public static string MonthSheetName(int month)
        {
            return month.ToString(CultureInfo.InvariantCulture);
        }

        public static string ExpectedTitle(string settlementKind)
        {
            if (settlementKind == GuangdongStage2SettlementKinds.Proxy)
            {
                return "代理费用结算单";
            }

            if (settlementKind == GuangdongStage2SettlementKinds.Intermediary)
            {
                return "居间费用结算单";
            }

            return "退补电费结算单";
        }

        public static string OutputFolderName(string settlementKind)
        {
            if (settlementKind == GuangdongStage2SettlementKinds.Proxy)
            {
                return "代理";
            }

            if (settlementKind == GuangdongStage2SettlementKinds.Intermediary)
            {
                return "居间";
            }

            return "退补";
        }

        public static string CreateUniqueRunDirectory(string outputRoot, int year, int month)
        {
            var stamp = DateTime.Now.ToString("yyyyMMdd-HHmmss", CultureInfo.InvariantCulture);
            var baseName = string.Format(
                CultureInfo.InvariantCulture,
                "广东{0}年{1}月分表初始化-{2}",
                year,
                month,
                stamp);
            var candidate = Path.Combine(outputRoot, baseName);
            var suffix = 2;
            while (Directory.Exists(candidate))
            {
                candidate = Path.Combine(outputRoot, baseName + "-" + suffix.ToString(CultureInfo.InvariantCulture));
                suffix++;
            }

            Directory.CreateDirectory(candidate);
            return candidate;
        }

        private static string AppendDirectorySeparatorChar(string path)
        {
            return path.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal)
                ? path
                : path + Path.DirectorySeparatorChar;
        }
    }
}
