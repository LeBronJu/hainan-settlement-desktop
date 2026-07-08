using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using HainanSettlementTool.Core.Models;
using HainanSettlementTool.Core.Services;

namespace HainanSettlementTool.Excel
{
    internal static class ChongqingStage2LedgerReader
    {
        public static List<ChongqingSettlementDetail> ReadDetails(ChongqingStage2Options options)
        {
            using (var stream = File.Open(options.LedgerPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var workbook = new XLWorkbook(stream))
            {
                var worksheet = FindLedgerWorksheet(workbook);
                var map = FindLedgerMap(worksheet, options.Month);
                return ReadDetails(worksheet, map);
            }
        }

        public static List<GroupSettlementTotal> BuildGroups(IList<ChongqingSettlementDetail> details)
        {
            return details
                .GroupBy(detail => ChongqingStage2Keys.SummaryKey(detail.Entity, detail.Kind))
                .Select(group =>
                {
                    var first = group.First();
                    return new GroupSettlementTotal
                    {
                        Kind = first.Kind,
                        Owner = first.Owner,
                        Entity = first.Entity,
                        DisplayEntity = first.Entity,
                        Rows = group.Count(),
                        ExpectedNet = Math.Round(group.Sum(item => item.ExpectedNet), 4)
                    };
                })
                .OrderBy(group => group.Kind)
                .ThenBy(group => group.Owner)
                .ThenBy(group => group.Entity)
                .ToList();
        }

        private static List<ChongqingSettlementDetail> ReadDetails(IXLWorksheet worksheet, ChongqingLedgerMap map)
        {
            var details = new List<ChongqingSettlementDetail>();
            var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? ChongqingStage2Layout.LedgerDataStartRow;
            for (var row = ChongqingStage2Layout.LedgerDataStartRow; row <= lastRow; row++)
            {
                var customer = ChongqingStage2ExcelUtil.CellText(worksheet.Cell(row, map.CustomerNameColumn));
                if (string.IsNullOrWhiteSpace(customer))
                {
                    continue;
                }

                var owner = ChongqingStage2ExcelUtil.CellText(worksheet.Cell(row, map.OwnerColumn));
                var proxyEntity = ChongqingStage2ExcelUtil.CellText(worksheet.Cell(row, map.ProjectDeveloperColumn));
                var agentOrSelf = ChongqingStage2ExcelUtil.CellText(worksheet.Cell(row, map.AgentOrSelfColumn));
                var intermediaryEntity = map.IntermediaryColumn > 0
                    ? ChongqingStage2ExcelUtil.CellText(worksheet.Cell(row, map.IntermediaryColumn))
                    : string.Empty;
                var refundEntity = map.PayeeColumn > 0
                    ? ChongqingStage2ExcelUtil.CellText(worksheet.Cell(row, map.PayeeColumn))
                    : string.Empty;

                AddDetailIfNeeded(
                    details,
                    CreateProxyLikeDetail(worksheet, map, row, customer, owner, proxyEntity, ChongqingStage2SettlementKinds.Proxy, map.ProxyRatioColumn, map.ProxyUnitPriceColumn, map.ProxyTaxRateColumn, map.ProxyNetColumn, map.RecoverShortfallColumn),
                    !TextUtil.S(agentOrSelf).Contains("自营"));
                AddDetailIfNeeded(
                    details,
                    CreateProxyLikeDetail(worksheet, map, row, customer, owner, intermediaryEntity, ChongqingStage2SettlementKinds.Intermediary, map.IntermediaryRatioColumn, map.IntermediaryUnitPriceColumn, map.IntermediaryTaxRateColumn, map.IntermediaryNetColumn, 0),
                    true);
                AddDetailIfNeeded(
                    details,
                    CreateRefundDetail(worksheet, map, row, customer, owner, refundEntity),
                    true);
            }

            return details;
        }

        private static ChongqingSettlementDetail CreateProxyLikeDetail(
            IXLWorksheet worksheet,
            ChongqingLedgerMap map,
            int row,
            string customer,
            string owner,
            string entity,
            string kind,
            int ratioColumn,
            int unitPriceColumn,
            int taxRateColumn,
            int netColumn,
            int recoverShortfallColumn)
        {
            var detail = CreateBaseDetail(worksheet, map, row, customer, owner, entity, kind);
            detail.Ratio = ChongqingStage2ExcelUtil.GetNumeric(worksheet, row, ratioColumn);
            detail.UnitPrice = ChongqingStage2ExcelUtil.GetNumeric(worksheet, row, unitPriceColumn);
            detail.TaxRate = ChongqingStage2ExcelUtil.GetNumeric(worksheet, row, taxRateColumn);
            detail.RecoverShortfall = recoverShortfallColumn > 0 ? ChongqingStage2ExcelUtil.GetNumeric(worksheet, row, recoverShortfallColumn) : 0d;
            detail.LedgerNet = ChongqingStage2ExcelUtil.GetNumeric(worksheet, row, netColumn);
            detail.Gross = Math.Round(detail.Total * detail.Ratio * detail.UnitPrice / 10d, 4);
            detail.AdjustedGross = Math.Round(detail.Gross - detail.RecoverShortfall, 4);
            detail.TaxAmount = Math.Round(detail.AdjustedGross / 1.13d * detail.TaxRate, 4);
            detail.CalculatedNet = Math.Round(detail.AdjustedGross - detail.TaxAmount, 4);
            detail.ExpectedNet = ChongqingStage2ExcelUtil.NonZeroOrFallback(detail.CalculatedNet, detail.LedgerNet);
            return detail;
        }

        private static ChongqingSettlementDetail CreateRefundDetail(
            IXLWorksheet worksheet,
            ChongqingLedgerMap map,
            int row,
            string customer,
            string owner,
            string entity)
        {
            var detail = CreateBaseDetail(worksheet, map, row, customer, owner, entity, ChongqingStage2SettlementKinds.Refund);
            detail.Ratio = ChongqingStage2ExcelUtil.GetNumeric(worksheet, row, map.RefundRatioColumn);
            detail.RefundSharpPrice = ChongqingStage2ExcelUtil.GetNumeric(worksheet, row, map.RefundSharpPriceColumn);
            detail.RefundPeakPrice = ChongqingStage2ExcelUtil.GetNumeric(worksheet, row, map.RefundPeakPriceColumn);
            detail.RefundFlatPrice = ChongqingStage2ExcelUtil.GetNumeric(worksheet, row, map.RefundFlatPriceColumn);
            detail.RefundValleyPrice = ChongqingStage2ExcelUtil.GetNumeric(worksheet, row, map.RefundValleyPriceColumn);
            detail.TaxRate = ChongqingStage2ExcelUtil.GetNumeric(worksheet, row, map.RefundTaxRateColumn);
            detail.LedgerNet = ChongqingStage2ExcelUtil.GetNumeric(worksheet, row, map.RefundNetColumn);
            detail.Gross = Math.Round(
                (detail.Sharp * detail.Ratio * detail.RefundSharpPrice
                    + detail.Peak * detail.Ratio * detail.RefundPeakPrice
                    + detail.Flat * detail.Ratio * detail.RefundFlatPrice
                    + detail.Valley * detail.Ratio * detail.RefundValleyPrice) / 10d,
                4);
            detail.AdjustedGross = detail.Gross;
            detail.TaxAmount = Math.Round(detail.Gross / 1.13d * detail.TaxRate, 4);
            detail.CalculatedNet = Math.Round(detail.Gross - detail.TaxAmount, 4);
            detail.ExpectedNet = ChongqingStage2ExcelUtil.NonZeroOrFallback(detail.CalculatedNet, detail.LedgerNet);
            return detail;
        }

        private static ChongqingSettlementDetail CreateBaseDetail(
            IXLWorksheet worksheet,
            ChongqingLedgerMap map,
            int row,
            string customer,
            string owner,
            string entity,
            string kind)
        {
            return new ChongqingSettlementDetail
            {
                LedgerRow = row,
                Customer = customer,
                Owner = owner,
                Entity = entity,
                Kind = kind,
                Total = ChongqingStage2ExcelUtil.GetNumeric(worksheet, row, map.TotalPowerColumn),
                Sharp = ChongqingStage2ExcelUtil.GetNumeric(worksheet, row, map.SharpPowerColumn),
                Peak = ChongqingStage2ExcelUtil.GetNumeric(worksheet, row, map.PeakPowerColumn),
                Flat = ChongqingStage2ExcelUtil.GetNumeric(worksheet, row, map.FlatPowerColumn),
                Valley = ChongqingStage2ExcelUtil.GetNumeric(worksheet, row, map.ValleyPowerColumn)
            };
        }

        private static void AddDetailIfNeeded(IList<ChongqingSettlementDetail> details, ChongqingSettlementDetail detail, bool canSettle)
        {
            if (!canSettle
                || detail == null
                || string.IsNullOrWhiteSpace(detail.Entity)
                || (Math.Abs(detail.LedgerNet) <= Stage2SettlementCalculator.AmountTolerance
                    && Math.Abs(detail.CalculatedNet) <= Stage2SettlementCalculator.AmountTolerance))
            {
                return;
            }

            details.Add(detail);
        }

        private static IXLWorksheet FindLedgerWorksheet(XLWorkbook workbook)
        {
            foreach (var worksheet in workbook.Worksheets)
            {
                if (FindHeaderColumn(worksheet, "电力用户名称") > 0)
                {
                    return worksheet;
                }
            }

            throw new InvalidOperationException("重庆台账中未找到表头“电力用户名称”。");
        }

        private static ChongqingLedgerMap FindLedgerMap(IXLWorksheet worksheet, int month)
        {
            var start = FindMonthStartColumn(worksheet, month);
            if (ChongqingStage2ExcelUtil.CellText(worksheet.Cell(2, start)) != "总实际电量（兆瓦时）"
                || ChongqingStage2ExcelUtil.CellText(worksheet.Cell(2, start + 1)) != "实际电量（兆瓦时）"
                || ChongqingStage2ExcelUtil.CellText(worksheet.Cell(3, start + 1)) != "尖"
                || ChongqingStage2ExcelUtil.CellText(worksheet.Cell(3, start + 2)) != "峰"
                || ChongqingStage2ExcelUtil.CellText(worksheet.Cell(3, start + 3)) != "平"
                || ChongqingStage2ExcelUtil.CellText(worksheet.Cell(3, start + 4)) != "谷")
            {
                throw new InvalidOperationException("重庆台账" + month + "月月度区块表头不符合预期。");
            }

            return new ChongqingLedgerMap
            {
                CustomerNameColumn = RequireHeaderColumn(worksheet, "电力用户名称"),
                ProjectDeveloperColumn = RequireHeaderColumn(worksheet, "项目开发人"),
                AgentOrSelfColumn = RequireHeaderColumn(worksheet, "代理或自营"),
                OwnerColumn = RequireHeaderColumn(worksheet, "负责人"),
                IntermediaryColumn = FindHeaderColumn(worksheet, "居间人"),
                PayeeColumn = FindHeaderColumn(worksheet, "收款人"),
                TotalPowerColumn = start,
                SharpPowerColumn = start + 1,
                PeakPowerColumn = start + 2,
                FlatPowerColumn = start + 3,
                ValleyPowerColumn = start + 4,
                IntermediaryRatioColumn = start + 7,
                IntermediaryUnitPriceColumn = start + 8,
                IntermediaryTaxRateColumn = start + 10,
                IntermediaryNetColumn = start + 12,
                RefundRatioColumn = start + 13,
                RefundSharpPriceColumn = start + 14,
                RefundPeakPriceColumn = start + 15,
                RefundFlatPriceColumn = start + 16,
                RefundValleyPriceColumn = start + 17,
                RefundTaxRateColumn = start + 19,
                RefundNetColumn = start + 21,
                ProxyRatioColumn = start + 22,
                ProxyUnitPriceColumn = start + 23,
                ProxyTaxRateColumn = start + 25,
                ProxyNetColumn = start + 27,
                RecoverShortfallColumn = start + 28
            };
        }

        private static int FindMonthStartColumn(IXLWorksheet worksheet, int month)
        {
            var label = month.ToString(CultureInfo.InvariantCulture) + "月";
            var lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? 1;
            for (var column = 1; column <= lastColumn; column++)
            {
                if (ChongqingStage2ExcelUtil.CellText(worksheet.Cell(1, column)) == label)
                {
                    return column;
                }
            }

            throw new InvalidOperationException("重庆台账中未找到" + label + "月度区块。");
        }

        private static int RequireHeaderColumn(IXLWorksheet worksheet, string headerText)
        {
            var column = FindHeaderColumn(worksheet, headerText);
            if (column <= 0)
            {
                throw new InvalidOperationException("重庆台账中未找到表头“" + headerText + "”。");
            }

            return column;
        }

        private static int FindHeaderColumn(IXLWorksheet worksheet, string headerText)
        {
            var lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? 1;
            for (var row = 1; row <= Math.Min(10, worksheet.LastRowUsed()?.RowNumber() ?? 1); row++)
            {
                for (var column = 1; column <= lastColumn; column++)
                {
                    if (ChongqingStage2ExcelUtil.CellText(worksheet.Cell(row, column)) == headerText)
                    {
                        return column;
                    }
                }
            }

            return 0;
        }
    }
}
