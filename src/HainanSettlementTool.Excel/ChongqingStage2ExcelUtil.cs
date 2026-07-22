using System;
using System.Globalization;
using System.IO;
using ClosedXML.Excel;
using HainanSettlementTool.Core.Models;
using HainanSettlementTool.Core.Services;

namespace HainanSettlementTool.Excel
{
    internal static class ChongqingStage2ExcelUtil
    {
        public const double TaxRateTolerance = 0.0000000001d;

        public static bool TaxRatesEqual(double left, double right)
        {
            return Math.Abs(left - right) <= TaxRateTolerance;
        }

        public static double GetNumeric(IXLWorksheet worksheet, int row, int column)
        {
            return ClosedXmlUtil.CellNumber(worksheet.Cell(row, column));
        }

        public static string CellText(IXLCell cell)
        {
            return TextUtil.S(cell.GetFormattedString());
        }

        public static double NonZeroOrFallback(double preferred, double fallback)
        {
            return Math.Abs(preferred) > Stage2SettlementCalculator.AmountTolerance ? preferred : fallback;
        }

        public static void SaveWorkbook(XLWorkbook workbook, string outputPath)
        {
            workbook.CalculateMode = XLCalculateMode.Auto;
            workbook.SaveAs(outputPath, new SaveOptions { EvaluateFormulasBeforeSaving = true });
        }

        public static XLWorkbook OpenWorkbookShared(string path)
        {
            using (var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                return new XLWorkbook(stream);
            }
        }

        public static void CopyWorkbookShared(string source, string target, bool overwrite)
        {
            if (File.Exists(target))
            {
                if (!overwrite)
                {
                    return;
                }

                File.Delete(target);
            }

            Directory.CreateDirectory(Path.GetDirectoryName(target));
            using (var input = File.Open(source, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var output = File.Create(target))
            {
                input.CopyTo(output);
            }
        }

        public static string TemplateRootFor(ChongqingStage2Options options, string kind)
        {
            if (kind == ChongqingStage2SettlementKinds.Proxy)
            {
                return options.ProxyTemplateDirectory;
            }

            if (kind == ChongqingStage2SettlementKinds.Intermediary)
            {
                return options.IntermediaryTemplateDirectory;
            }

            return options.RefundTemplateDirectory;
        }

        public static string OutputRootFor(ChongqingStage2Options options, string kind)
        {
            return Path.Combine(options.OutputDirectory, OutputFolderName(kind));
        }

        public static string OutputFolderName(string kind)
        {
            if (kind == ChongqingStage2SettlementKinds.Proxy)
            {
                return "2026年代理 - 重庆";
            }

            if (kind == ChongqingStage2SettlementKinds.Intermediary)
            {
                return "2026年居间 - 重庆";
            }

            return "2026年退补 - 重庆";
        }

        public static string KindShort(string kind)
        {
            if (kind == ChongqingStage2SettlementKinds.Proxy)
            {
                return "代理";
            }

            if (kind == ChongqingStage2SettlementKinds.Intermediary)
            {
                return "居间";
            }

            return "退补";
        }

        public static string RelativePath(string root, string path)
        {
            var rootUri = new Uri(AppendDirectorySeparatorChar(Path.GetFullPath(root)));
            var pathUri = new Uri(Path.GetFullPath(path));
            return Uri.UnescapeDataString(rootUri.MakeRelativeUri(pathUri).ToString()).Replace('/', Path.DirectorySeparatorChar);
        }

        public static string MonthSheetName(int month)
        {
            return month.ToString(CultureInfo.InvariantCulture);
        }

        private static string AppendDirectorySeparatorChar(string path)
        {
            return path.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal) ? path : path + Path.DirectorySeparatorChar;
        }
    }
}
