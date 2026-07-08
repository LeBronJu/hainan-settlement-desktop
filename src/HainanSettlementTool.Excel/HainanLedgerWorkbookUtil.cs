using System;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using HainanSettlementTool.Core.Services;

namespace HainanSettlementTool.Excel
{
    internal static class HainanLedgerWorkbookUtil
    {
        public static IXLWorksheet MainSheet(XLWorkbook workbook)
        {
            var named = workbook.Worksheets.FirstOrDefault(ws => ws.Name == HainanLedgerLayout.MainSheetName);
            if (named != null)
            {
                return named;
            }

            var sheet1 = workbook.Worksheets.FirstOrDefault(ws => ws.Name == "Sheet1");
            if (sheet1 != null)
            {
                return sheet1;
            }

            var matched = workbook.Worksheets.FirstOrDefault(ws => TextUtil.S(ws.Cell("A1").GetFormattedString()).Contains("售电结算台账"));
            if (matched != null)
            {
                return matched;
            }

            throw new InvalidOperationException("找不到海南台账主表：" + HainanLedgerLayout.MainSheetName);
        }

        public static string DefaultStage1OutputLedgerName(string baseLedger, int month)
        {
            var stem = Path.GetFileNameWithoutExtension(baseLedger);
            var extension = Path.GetExtension(baseLedger);
            return stem + "】补" + month + "月电量" + extension;
        }
    }
}
