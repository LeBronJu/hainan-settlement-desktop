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
            return ReadSnapshot(options).Details;
        }

        public static ChongqingStage2LedgerSnapshot ReadSnapshot(ChongqingStage2Options options)
        {
            using (var stream = File.Open(options.LedgerPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var workbook = new XLWorkbook(stream))
            {
                var worksheet = FindLedgerWorksheet(workbook);
                var map = FindLedgerMap(worksheet, options.Month);
                return ReadSnapshot(worksheet, map);
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

        private static ChongqingStage2LedgerSnapshot ReadSnapshot(IXLWorksheet worksheet, ChongqingLedgerMap map)
        {
            var snapshot = new ChongqingStage2LedgerSnapshot();
            var occurrences = new List<ChongqingRelationshipOccurrence>();
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
                var isSelfOperated = TextUtil.S(agentOrSelf).Contains("自营");

                if (isSelfOperated)
                {
                    ValidateSelfOperatedProxyFieldsEmpty(
                        snapshot,
                        worksheet,
                        row,
                        customer,
                        owner,
                        proxyEntity,
                        map.ProxyRatioColumn,
                        map.ProxyUnitPriceColumn,
                        map.ProxyTaxRateColumn);
                }
                else
                {
                    ValidateProxyLikeRelationship(
                        snapshot,
                        occurrences,
                        worksheet,
                        row,
                        customer,
                        owner,
                        proxyEntity,
                        ChongqingStage2SettlementKinds.Proxy,
                        map.ProxyRatioColumn,
                        map.ProxyUnitPriceColumn,
                        map.ProxyTaxRateColumn);
                }

                ValidateProxyLikeRelationship(
                    snapshot,
                    occurrences,
                    worksheet,
                    row,
                    customer,
                    owner,
                    intermediaryEntity,
                    ChongqingStage2SettlementKinds.Intermediary,
                    map.IntermediaryRatioColumn,
                    map.IntermediaryUnitPriceColumn,
                    map.IntermediaryTaxRateColumn);

                AddRefundOccurrence(
                    occurrences,
                    worksheet,
                    map,
                    row,
                    customer,
                    owner,
                    refundEntity);

                AddDetailIfNeeded(
                    snapshot.Details,
                    isSelfOperated
                        ? null
                        : CreateProxyLikeDetail(worksheet, map, row, customer, owner, proxyEntity, ChongqingStage2SettlementKinds.Proxy, map.ProxyRatioColumn, map.ProxyUnitPriceColumn, map.ProxyTaxRateColumn, map.ProxyNetColumn, map.RecoverShortfallColumn),
                    !isSelfOperated);
                AddDetailIfNeeded(
                    snapshot.Details,
                    CreateProxyLikeDetail(worksheet, map, row, customer, owner, intermediaryEntity, ChongqingStage2SettlementKinds.Intermediary, map.IntermediaryRatioColumn, map.IntermediaryUnitPriceColumn, map.IntermediaryTaxRateColumn, map.IntermediaryNetColumn, 0),
                    true);
                AddDetailIfNeeded(
                    snapshot.Details,
                    CreateRefundDetail(worksheet, map, row, customer, owner, refundEntity),
                    true);
            }

            ApplyCanonicalOwnersAndGroupChecks(snapshot, occurrences);
            return snapshot;
        }

        private static void ValidateSelfOperatedProxyFieldsEmpty(
            ChongqingStage2LedgerSnapshot snapshot,
            IXLWorksheet worksheet,
            int row,
            string customer,
            string owner,
            string entity,
            int ratioColumn,
            int unitPriceColumn,
            int taxRateColumn)
        {
            var ratio = ReadParameterValue(worksheet.Cell(row, ratioColumn));
            var unitPrice = ReadParameterValue(worksheet.Cell(row, unitPriceColumn));
            var taxRate = ReadParameterValue(worksheet.Cell(row, taxRateColumn));
            if (string.IsNullOrWhiteSpace(entity)
                && !ratio.HasContent
                && !unitPrice.HasContent
                && !taxRate.HasContent)
            {
                return;
            }

            var residuals = new List<string>();
            if (!string.IsNullOrWhiteSpace(entity))
            {
                residuals.Add("代理主体：" + entity);
            }

            AddResidualParameter(residuals, "比例", ratio);
            AddResidualParameter(residuals, "单价", unitPrice);
            AddResidualParameter(residuals, "税率", taxRate);
            snapshot.Issues.Add(new ChongqingStage2CheckIssue
            {
                Code = Stage2PreflightIssueKinds.RelationshipParametersInvalid,
                Disposition = Stage2PreflightDisposition.Blocker,
                Severity = "阻断",
                Category = "自营行残留代理关系字段",
                Kind = Stage2PreflightIssueKinds.RelationshipParametersInvalid,
                SettlementKind = ChongqingStage2SettlementKinds.Proxy,
                Customer = customer,
                Owner = owner,
                Entity = entity,
                LedgerRow = row,
                CurrentValue = string.Join("；", residuals),
                Message = "重庆台账第" + row + "行为自营，但仍填写了代理主体或代理参数。",
                Suggestion = "自营行的代理主体、比例、单价和税率必须全部空白；请清理残留内容后重新预检。"
            });
        }

        private static void AddResidualParameter(
            IList<string> residuals,
            string name,
            Stage2RelationshipParameterValue value)
        {
            if (!value.HasContent)
            {
                return;
            }

            residuals.Add(name + "：" + (string.IsNullOrWhiteSpace(value.DisplayValue)
                ? "（公式或不可见内容）"
                : value.DisplayValue));
        }

        private static void AddRefundOccurrence(
            IList<ChongqingRelationshipOccurrence> occurrences,
            IXLWorksheet worksheet,
            ChongqingLedgerMap map,
            int row,
            string customer,
            string owner,
            string entity)
        {
            if (string.IsNullOrWhiteSpace(entity))
            {
                return;
            }

            // 重庆退补使用分段单价结构。当前没有业务依据要求四个分段
            // 单价必须大于零，因此这里只投影已能可靠识别的主体、负责人
            // 和扣税率，供跨费用类型一致的聚合检查使用。
            occurrences.Add(new ChongqingRelationshipOccurrence
            {
                LedgerRow = row,
                Customer = customer,
                Owner = owner,
                Entity = entity,
                Kind = ChongqingStage2SettlementKinds.Refund,
                TaxRate = ChongqingStage2ExcelUtil.GetNumeric(
                    worksheet,
                    row,
                    map.RefundTaxRateColumn)
            });
        }

        private static void ValidateProxyLikeRelationship(
            ChongqingStage2LedgerSnapshot snapshot,
            IList<ChongqingRelationshipOccurrence> occurrences,
            IXLWorksheet worksheet,
            int row,
            string customer,
            string owner,
            string entity,
            string kind,
            int ratioColumn,
            int unitPriceColumn,
            int taxRateColumn)
        {
            var ratio = ReadParameterValue(worksheet.Cell(row, ratioColumn));
            var unitPrice = ReadParameterValue(worksheet.Cell(row, unitPriceColumn));
            var taxRate = ReadParameterValue(worksheet.Cell(row, taxRateColumn));
            var validation = Stage2RelationshipParameterValidator.Validate(entity, ratio, unitPrice, taxRate);
            foreach (var error in validation.Errors)
            {
                var parametersWithoutSubject = error.Kind == Stage2RelationshipParameterErrorKind.ParametersWithoutSubject;
                snapshot.Issues.Add(new ChongqingStage2CheckIssue
                {
                    Code = parametersWithoutSubject
                        ? Stage2PreflightIssueKinds.RelationshipParametersWithoutSubject
                        : Stage2PreflightIssueKinds.RelationshipParametersInvalid,
                    Disposition = Stage2PreflightDisposition.Blocker,
                    Severity = "阻断",
                    Category = parametersWithoutSubject ? "主体为空但参数有值" : "结算关系参数不完整",
                    Kind = parametersWithoutSubject
                        ? Stage2PreflightIssueKinds.RelationshipParametersWithoutSubject
                        : Stage2PreflightIssueKinds.RelationshipParametersInvalid,
                    SettlementKind = kind,
                    Customer = customer,
                    Owner = owner,
                    Entity = entity,
                    LedgerRow = row,
                    CurrentValue = error.ParameterName + "：" + TextUtil.S(error.DisplayValue),
                    Message = "重庆台账第" + row + "行" + kind + "关系的" + error.ParameterName + "填写不符合规则。",
                    Suggestion = parametersWithoutSubject
                        ? "主体为空时比例、单价和税率必须全部空白。"
                        : "主体已填写时比例、单价和税率必须全部为大于 0 的数字。"
                });
            }

            if (!validation.HasRelationship || !validation.IsValid)
            {
                return;
            }

            occurrences.Add(new ChongqingRelationshipOccurrence
            {
                LedgerRow = row,
                Customer = customer,
                Owner = owner,
                Entity = entity,
                Kind = kind,
                TaxRate = taxRate.Value
            });
        }

        private static Stage2RelationshipParameterValue ReadParameterValue(IXLCell cell)
        {
            var hasFormula = !string.IsNullOrWhiteSpace(cell.FormulaA1);
            var formatted = cell.GetFormattedString();
            var hasContent = hasFormula || !string.IsNullOrWhiteSpace(formatted);
            var value = 0d;
            var numeric = hasContent && cell.TryGetValue(out value);
            return new Stage2RelationshipParameterValue
            {
                HasContent = hasContent,
                IsNumeric = numeric,
                Value = numeric ? value : 0d,
                DisplayValue = formatted
            };
        }

        private static void ApplyCanonicalOwnersAndGroupChecks(
            ChongqingStage2LedgerSnapshot snapshot,
            IList<ChongqingRelationshipOccurrence> occurrences)
        {
            foreach (var group in occurrences
                .GroupBy(item => ChongqingStage2Keys.SummaryKey(item.Entity, item.Kind)))
            {
                var ordered = group.OrderBy(item => item.LedgerRow).ToList();
                var first = ordered[0];
                if (string.IsNullOrWhiteSpace(first.Owner))
                {
                    snapshot.Issues.Add(new ChongqingStage2CheckIssue
                    {
                        Code = Stage2PreflightIssueKinds.FirstOwnerMissing,
                        Disposition = Stage2PreflightDisposition.Blocker,
                        Severity = "阻断",
                        Category = "首次关系行负责人为空",
                        Kind = Stage2PreflightIssueKinds.FirstOwnerMissing,
                        SettlementKind = first.Kind,
                        Customer = first.Customer,
                        Entity = first.Entity,
                        LedgerRow = first.LedgerRow,
                        Message = "重庆台账中该主体的首次关系行没有负责人，无法确定唯一分表目录。",
                        Suggestion = "请补齐第" + first.LedgerRow + "行负责人后重新预检。"
                    });
                }

                var owners = ordered
                    .Select(item => TextUtil.S(item.Owner))
                    .Where(item => !string.IsNullOrWhiteSpace(item))
                    .Distinct()
                    .ToList();
                if (owners.Count > 1)
                {
                    snapshot.Issues.Add(new ChongqingStage2CheckIssue
                    {
                        Code = Stage2PreflightIssueKinds.MultipleOwners,
                        Disposition = Stage2PreflightDisposition.Review,
                        Severity = "复核",
                        Category = "同一主体存在多个负责人",
                        Kind = Stage2PreflightIssueKinds.MultipleOwners,
                        SettlementKind = first.Kind,
                        Owner = first.Owner,
                        Entity = first.Entity,
                        LedgerRow = first.LedgerRow,
                        PreviousValue = string.Join("、", owners),
                        CurrentValue = "本次归属：" + first.Owner,
                        Message = "程序将合并该主体在多个负责人名下的客户和金额，只生成一份分表及一条汇总记录。",
                        Suggestion = "请确认按首次关系行负责人“" + first.Owner + "”归档。"
                    });
                }

                foreach (var detail in snapshot.Details.Where(item =>
                    ChongqingStage2Keys.SummaryKey(item.Entity, item.Kind) == group.Key))
                {
                    detail.Owner = first.Owner;
                }

                var relationshipTaxRates = ordered
                    .Select(item => item.TaxRate)
                    .Distinct(new TaxRateComparer())
                    .ToList();
                if (relationshipTaxRates.Count > 1)
                {
                    snapshot.Issues.Add(new ChongqingStage2CheckIssue
                    {
                        Code = Stage2PreflightIssueKinds.ConflictingTaxRates,
                        Disposition = Stage2PreflightDisposition.Blocker,
                        Severity = "阻断",
                        Category = "同一汇总主体扣税率冲突",
                        Kind = Stage2PreflightIssueKinds.ConflictingTaxRates,
                        SettlementKind = first.Kind,
                        Owner = first.Owner,
                        Entity = first.Entity,
                        LedgerRow = first.LedgerRow,
                    CurrentValue = string.Join("、", relationshipTaxRates.Select(item => item.ToString("0.##########", CultureInfo.InvariantCulture))),
                        Message = "同一费用类型和主体出现多个扣税率，单条汇总记录无法可靠表达。",
                        Suggestion = "请检查台账税率并统一后重新预检。"
                    });
                }
            }

            AddActiveTaxConflictIssues(snapshot);
        }

        private static void AddActiveTaxConflictIssues(ChongqingStage2LedgerSnapshot snapshot)
        {
            foreach (var group in snapshot.Details
                .GroupBy(item => ChongqingStage2Keys.SummaryKey(item.Entity, item.Kind)))
            {
                var first = group.OrderBy(item => item.LedgerRow).First();
                var rates = group
                    .Select(item => item.TaxRate)
                    .Distinct(new TaxRateComparer())
                    .ToList();
                if (rates.Count <= 1 || snapshot.Issues.Any(issue =>
                    issue.Code == Stage2PreflightIssueKinds.ConflictingTaxRates
                    && ChongqingStage2Keys.SummaryKey(issue.Entity, issue.SettlementKind) == group.Key))
                {
                    continue;
                }

                snapshot.Issues.Add(new ChongqingStage2CheckIssue
                {
                    Code = Stage2PreflightIssueKinds.ConflictingTaxRates,
                    Disposition = Stage2PreflightDisposition.Blocker,
                    Severity = "阻断",
                    Category = "同一汇总主体扣税率冲突",
                    Kind = Stage2PreflightIssueKinds.ConflictingTaxRates,
                    SettlementKind = first.Kind,
                    Owner = first.Owner,
                    Entity = first.Entity,
                    LedgerRow = first.LedgerRow,
                    CurrentValue = string.Join("、", rates.Select(item => item.ToString("0.##########", CultureInfo.InvariantCulture))),
                    Message = "同一费用类型和主体出现多个扣税率，单条汇总记录无法可靠表达。",
                    Suggestion = "请检查台账税率并统一后重新预检。"
                });
            }
        }

        private sealed class TaxRateComparer : IEqualityComparer<double>
        {
            public bool Equals(double x, double y)
            {
                return ChongqingStage2ExcelUtil.TaxRatesEqual(x, y);
            }

            public int GetHashCode(double obj)
            {
                return 0;
            }
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
            detail.ExpectedNet = detail.CalculatedNet;
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
            detail.ExpectedNet = detail.CalculatedNet;
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
