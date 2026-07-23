using System;
using ClosedXML.Excel;
using HainanSettlementTool.Core.Services;

namespace HainanSettlementTool.Excel
{
    internal static class ClosedXmlUtil
    {
        public static double CellNumber(IXLCell cell)
        {
            if (cell.DataType == XLDataType.Number)
            {
                return cell.GetDouble();
            }

            return TextUtil.N(cell.GetFormattedString());
        }

        public static string ColumnLetter(int columnNumber)
        {
            var dividend = columnNumber;
            var columnName = string.Empty;
            while (dividend > 0)
            {
                var modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }

            return columnName;
        }

        public static void SetOnlyActiveWorksheet(
            XLWorkbook workbook,
            IXLWorksheet worksheet)
        {
            worksheet.Visibility = XLWorksheetVisibility.Visible;
            foreach (var candidate in workbook.Worksheets)
            {
                candidate.SetTabSelected(false);
            }

            worksheet.SetTabSelected();
            worksheet.SetTabActive();
        }
    }
}
