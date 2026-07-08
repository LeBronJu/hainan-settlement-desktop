using System;
using System.Collections.Generic;
using ClosedXML.Excel;
using HainanSettlementTool.Core.Models;
using HainanSettlementTool.Core.Services;

namespace HainanSettlementTool.Excel
{
    internal static class HainanStage2LedgerReader
    {
        internal static void ReadLedgerRows(string ledgerPath, int month, List<HainanStage2DetailSettlementRow> proxyRows, List<HainanStage2DetailSettlementRow> interRows)
        {
            using (var workbook = new XLWorkbook(ledgerPath))
            {
                var worksheet = HainanLedgerWorkbookUtil.MainSheet(workbook);
                var start = FindMonthStartColumn(worksheet, month);
                var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 1;

                for (var row = 4; row <= lastRow; row++)
                {
                    var customer = TextUtil.S(worksheet.Cell(row, 3).GetFormattedString());
                    if (string.IsNullOrWhiteSpace(customer))
                    {
                        continue;
                    }

                    var total = HainanStage2ExcelUtil.GetNumeric(worksheet, row, start);
                    if (total <= 0)
                    {
                        continue;
                    }

                    var owner = TextUtil.S(worksheet.Cell(row, 10).GetFormattedString());
                    var developer = TextUtil.S(worksheet.Cell(row, 8).GetFormattedString());
                    var interName = TextUtil.S(worksheet.Cell(row, 19).GetFormattedString());

                    var interRow = CreateDetailRow(worksheet, row, start, customer, owner, interName, "居间", start + 7, start + 8, start + 10, start + 12);
                    if (!string.IsNullOrWhiteSpace(interName) && HasSettlementAmount(interRow))
                    {
                        interRows.Add(interRow);
                    }

                    var proxyRow = CreateDetailRow(worksheet, row, start, customer, owner, developer, "代理", start + 13, start + 14, start + 16, start + 18);
                    if (!string.IsNullOrWhiteSpace(developer) && HasSettlementAmount(proxyRow))
                    {
                        proxyRows.Add(proxyRow);
                    }
                }
            }
        }

        private static HainanStage2DetailSettlementRow CreateDetailRow(
            IXLWorksheet worksheet,
            int ledgerRow,
            int start,
            string customer,
            string owner,
            string entity,
            string kind,
            int ratioColumn,
            int unitPriceColumn,
            int taxRateColumn,
            int cachedNetColumn)
        {
            var total = HainanStage2ExcelUtil.GetNumeric(worksheet, ledgerRow, start);
            var ratio = HainanStage2ExcelUtil.GetNumeric(worksheet, ledgerRow, ratioColumn);
            var unitPrice = HainanStage2ExcelUtil.GetNumeric(worksheet, ledgerRow, unitPriceColumn);
            var taxRate = HainanStage2ExcelUtil.GetNumeric(worksheet, ledgerRow, taxRateColumn);
            var ledgerNet = HainanStage2ExcelUtil.GetNumeric(worksheet, ledgerRow, cachedNetColumn);
            var amounts = Stage2SettlementCalculator.CalculateAmounts(total, ratio, unitPrice, taxRate);

            return new HainanStage2DetailSettlementRow
            {
                LedgerRow = ledgerRow,
                Customer = customer,
                Owner = owner,
                Entity = entity,
                Kind = kind,
                Total = total,
                Sharp = HainanStage2ExcelUtil.GetNumeric(worksheet, ledgerRow, start + 1),
                Peak = HainanStage2ExcelUtil.GetNumeric(worksheet, ledgerRow, start + 2),
                Flat = HainanStage2ExcelUtil.GetNumeric(worksheet, ledgerRow, start + 3),
                Valley = HainanStage2ExcelUtil.GetNumeric(worksheet, ledgerRow, start + 4),
                PeakFlat = HainanStage2ExcelUtil.GetNumeric(worksheet, ledgerRow, start + 5),
                ValleyFlat = HainanStage2ExcelUtil.GetNumeric(worksheet, ledgerRow, start + 6),
                Ratio = ratio,
                UnitPrice = unitPrice,
                TaxRate = taxRate,
                LedgerNet = ledgerNet,
                Gross = amounts.Gross,
                Adjustment = amounts.Adjustment,
                AdjustedGross = amounts.AdjustedGross,
                TaxAmount = amounts.TaxAmount,
                CalculatedNet = amounts.CalculatedNet,
                ExpectedNet = amounts.ExpectedNet
            };
        }

        private static bool HasSettlementAmount(HainanStage2DetailSettlementRow row)
        {
            return Math.Abs(row.LedgerNet) > Stage2SettlementCalculator.AmountTolerance
                || Math.Abs(row.CalculatedNet) > Stage2SettlementCalculator.AmountTolerance;
        }

        private static int FindMonthStartColumn(IXLWorksheet worksheet, int month)
        {
            var lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? 1;
            for (var column = 1; column <= lastColumn; column++)
            {
                if (TextUtil.S(worksheet.Cell(1, column).GetFormattedString()) == month + "月")
                {
                    return column;
                }
            }

            throw new InvalidOperationException("未找到 " + month + "月 的台账区块。");
        }
    }
}
