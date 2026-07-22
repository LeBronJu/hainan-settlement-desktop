using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using HainanSettlementTool.Core.Models;
using HainanSettlementTool.Core.Services;

namespace HainanSettlementTool.Excel
{
    internal static class ChongqingStage2SummaryWorkbookWriter
    {
        public static void VerifyGeneratedSummary(
            ChongqingStage2Options options,
            IList<GroupSettlementTotal> groups,
            IList<ChongqingSettlementDetail> details,
            string summaryPath)
        {
            using (var stream = File.Open(summaryPath, FileMode.Open, FileAccess.Read, FileShare.Read))
            using (var workbook = new XLWorkbook(stream))
            {
                var mainSheet = FindSummaryWorksheet(workbook);
                var monthColumns = FindSummaryMonthBlocks(mainSheet, options.Month);
                if (monthColumns.Count != 1)
                {
                    throw new InvalidDataException(
                        "重庆阶段二主汇总表的本次目标月份块数量不是 1："
                        + monthColumns.Count + "。");
                }

                var monthColumn = monthColumns[0];

                var mainRows = ReadSummaryMeta(mainSheet);
                var taxRateByKey = BuildTaxRateIndex(details);
                foreach (var group in groups)
                {
                    var key = ChongqingStage2Keys.SummaryKey(group.Entity, group.Kind);
                    var matches = mainRows
                        .Where(meta => ChongqingStage2Keys.SummaryKey(meta.Entity, meta.Kind) == key)
                        .ToList();
                    if (matches.Count != 1)
                    {
                        throw new InvalidDataException("重庆阶段二汇总表中主体记录数不是 1：" + group.Kind + " " + group.Entity);
                    }

                    var summaryRow = matches[0];
                    if (!ChongqingStage2PaymentParties.Supported.Contains(summaryRow.PaymentParty))
                    {
                        throw new InvalidDataException("重庆阶段二汇总表支付方不完整：" + group.Kind + " " + group.Entity);
                    }

                    var amount = ClosedXmlUtil.CellNumber(
                        mainSheet.Cell(summaryRow.Row, monthColumn + SummaryKindOffset(group.Kind)));
                    if (Math.Abs(amount - group.ExpectedNet) > Stage2SettlementCalculator.AmountTolerance)
                    {
                        throw new InvalidDataException("重庆阶段二汇总表本月金额与分组合计不一致：" + group.Kind + " " + group.Entity);
                    }

                    var expectedTaxRate = TaxRateFromIndex(taxRateByKey, group.Entity, group.Kind);
                    var withholdingTaxRate = ClosedXmlUtil.CellNumber(mainSheet.Cell(summaryRow.Row, 9));
                    var totalTaxRate = ClosedXmlUtil.CellNumber(mainSheet.Cell(summaryRow.Row, 10));
                    var expectedFormula = "J" + summaryRow.Row + "-I" + summaryRow.Row;
                    var actualOwner = ChongqingStage2ExcelUtil.CellText(mainSheet.Cell(summaryRow.Row, 11));
                    if (!ChongqingStage2ExcelUtil.TaxRatesEqual(withholdingTaxRate, expectedTaxRate)
                        || !ChongqingStage2ExcelUtil.TaxRatesEqual(totalTaxRate, 0.13d)
                        || TextUtil.CustomerKey(actualOwner) != TextUtil.CustomerKey(group.Owner)
                        || !string.Equals(
                            TextUtil.S(mainSheet.Cell(summaryRow.Row, 8).FormulaA1),
                            expectedFormula,
                            StringComparison.OrdinalIgnoreCase))
                    {
                        throw new InvalidDataException("重庆阶段二汇总表税率字段与台账规则不一致：" + group.Kind + " " + group.Entity);
                    }

                    VerifyPaymentPartyMembership(
                        workbook,
                        options.Month,
                        group,
                        summaryRow.PaymentParty,
                        summaryRow.Payee,
                        expectedTaxRate);
                }
            }
        }

        public static string BuildSummary(
            ChongqingStage2Options options,
            IList<GroupSettlementTotal> groups,
            IList<ChongqingSettlementDetail> details,
            IList<string> warnings)
        {
            var outputPath = PlanOutputPath(options);
            FileAccessGuard.RequireWritableWorkbook(outputPath, "重庆阶段二输出汇总表");
            ChongqingStage2ExcelUtil.CopyWorkbookShared(options.SummaryTemplatePath, outputPath, overwrite: true);

            using (var workbook = new XLWorkbook(outputPath))
            {
                var mainSheet = FindSummaryWorksheet(workbook);
                var mainMeta = ReadSummaryMeta(mainSheet);
                var paymentPartyByKey = BuildPaymentPartyIndex(options, groups, mainMeta);
                var taxRateByKey = BuildTaxRateIndex(details);
                var payeeByKey = BuildPayeeIndex(
                    groups,
                    mainMeta,
                    ReadReliableSummaryMeta(workbook, mainSheet, options.Month));
                var knownMainKeys = new HashSet<string>(mainMeta.Select(item =>
                    ChongqingStage2Keys.SummaryKey(item.Entity, item.Kind)));
                var newSubjectKeys = new HashSet<string>(groups
                    .Select(group => ChongqingStage2Keys.SummaryKey(group.Entity, group.Kind))
                    .Where(key => !knownMainKeys.Contains(key)));
                WriteSummarySheet(
                    mainSheet,
                    groups,
                    options.Month,
                    null,
                    warnings,
                    paymentPartyByKey,
                    taxRateByKey,
                    payeeByKey,
                    newSubjectKeys,
                    null);

                foreach (var paymentParty in ChongqingStage2PaymentParties.Supported)
                {
                    bool createdPaymentPartySheet;
                    var sheet = PreparePaymentPartySheet(workbook, paymentParty, options.Month, mainSheet, out createdPaymentPartySheet);
                    var allowedKeys = new HashSet<string>(paymentPartyByKey
                        .Where(item => item.Value == paymentParty)
                        .Select(item => item.Key));
                    var partyGroups = groups
                        .Where(group => PartyForSummaryTotal(group, paymentPartyByKey) == paymentParty)
                        .ToList();
                    foreach (var group in partyGroups)
                    {
                        allowedKeys.Add(ChongqingStage2Keys.SummaryKey(group.Entity, group.Kind));
                    }

                    WriteSummarySheet(
                        sheet,
                        partyGroups,
                        options.Month,
                        allowedKeys,
                        warnings,
                        paymentPartyByKey,
                        taxRateByKey,
                        payeeByKey,
                        newSubjectKeys,
                        mainSheet,
                        createdPaymentPartySheet);
                    ApplyPaymentPartyTitleMerge(sheet);
                }

                MoveTargetPaymentPartySheetsAfterSummary(workbook, options.Month, mainSheet);
                ChongqingStage2ExcelUtil.SaveWorkbook(workbook, outputPath);
            }

            return outputPath;
        }

        public static string PlanOutputPath(ChongqingStage2Options options)
        {
            var outputName = string.IsNullOrWhiteSpace(options.OutputSummaryName)
                ? "【2026年重庆代理费汇总表-" + options.Month + "月自动化】.xlsx"
                : options.OutputSummaryName;
            var outputRoot = Path.GetFullPath(options.OutputDirectory);
            var outputPath = Path.GetFullPath(Path.Combine(outputRoot, outputName));
            var prefix = outputRoot.TrimEnd(
                    Path.DirectorySeparatorChar,
                    Path.AltDirectorySeparatorChar)
                + Path.DirectorySeparatorChar;
            if (!outputPath.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(
                    "重庆阶段二输出汇总表路径必须严格位于本次输出目录内。" );
            }

            return outputPath;
        }

        public static void AddSummaryPaymentIssues(
            ChongqingStage2Options options,
            IList<GroupSettlementTotal> groups,
            IList<ChongqingStage2CheckIssue> issues)
        {
            using (var stream = File.Open(options.SummaryTemplatePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var workbook = new XLWorkbook(stream))
            {
                var mainSheet = FindSummaryWorksheet(workbook);
                var mainRows = ReadSummaryMeta(mainSheet);
                var sourceRows = ReadReliableSummaryMeta(workbook, mainSheet, options.Month);
                var activeByKey = groups.ToDictionary(
                    group => ChongqingStage2Keys.SummaryKey(group.Entity, group.Kind),
                    group => group);

                AddSummaryTargetMonthIssues(options, workbook, mainSheet, issues);
                AddDuplicateSummaryIssues(options, sourceRows, issues);

                var mainKeys = new HashSet<string>(mainRows.Select(item =>
                    ChongqingStage2Keys.SummaryKey(item.Entity, item.Kind)));
                foreach (var orphanGroup in sourceRows
                    .Where(item => !string.IsNullOrWhiteSpace(item.SourcePaymentParty))
                    .GroupBy(item => ChongqingStage2Keys.SummaryKey(item.Entity, item.Kind))
                    .Where(sourceGroup => !mainKeys.Contains(sourceGroup.Key)))
                {
                    var first = orphanGroup.First();
                    issues.Add(new ChongqingStage2CheckIssue
                    {
                        Code = Stage2PreflightIssueKinds.SummaryOrphanSubject,
                        Disposition = Stage2PreflightDisposition.Blocker,
                        Severity = "阻断",
                        Category = "支付方工作表存在主表外主体",
                        Kind = Stage2PreflightIssueKinds.SummaryOrphanSubject,
                        SettlementKind = first.Kind,
                        Entity = first.Entity,
                        TemplateFile = options.SummaryTemplatePath,
                        SheetName = string.Join("、", orphanGroup.Select(item => item.SheetName).Distinct()),
                        Message = "支付方工作表存在主汇总表没有的" + first.Kind + "主体“" + first.Entity + "”。",
                        Suggestion = "请先确认该行应补入主汇总表还是从支付方工作表移除，再重新预检。"
                    });
                }

                foreach (var sourceGroup in sourceRows
                    .GroupBy(item => ChongqingStage2Keys.SummaryKey(item.Entity, item.Kind)))
                {
                    AddPayeeSourceIssue(options, sourceGroup.ToList(), issues);
                    AddPaymentPartyConflictIssue(options, sourceGroup.ToList(), issues);
                }

                foreach (var mainGroup in mainRows
                    .GroupBy(item => ChongqingStage2Keys.SummaryKey(item.Entity, item.Kind)))
                {
                    if (mainGroup.Count() != 1)
                    {
                        continue;
                    }

                    var existing = mainGroup.First();
                    if (ChongqingStage2PaymentParties.Supported.Contains(TextUtil.S(existing.PaymentParty)))
                    {
                        continue;
                    }

                    GroupSettlementTotal active;
                    if (activeByKey.TryGetValue(mainGroup.Key, out active))
                    {
                        AddPaymentPartyRequirement(
                            options,
                            AvailablePaymentParties(workbook, options.Month),
                            mainSheet.Name,
                            active.Kind,
                            active.Entity,
                            active.Owner,
                            "存量汇总主体支付方补充",
                            "存量汇总主体支付方为空或不是清能/清辉，需要本次明确选择。",
                            issues);
                    }
                    else
                    {
                        issues.Add(new ChongqingStage2CheckIssue
                        {
                            Code = Stage2PreflightIssueKinds.PaymentPartyRequired,
                            Disposition = Stage2PreflightDisposition.Blocker,
                            Severity = "阻断",
                            Category = "非本月主体支付方缺失或不支持",
                            Kind = Stage2PreflightIssueKinds.PaymentPartyRequired,
                            SettlementKind = existing.Kind,
                            Entity = existing.Entity,
                            TemplateFile = options.SummaryTemplatePath,
                            SheetName = mainSheet.Name,
                            PreviousValue = "模板支付方：" + TextUtil.S(existing.PaymentParty),
                            Message = "主汇总表中的非本月主体支付方为空或不是清能/清辉，程序不能在没有本月业务决定的情况下重排历史支付方行。",
                            Suggestion = "请先在主汇总表中补齐正确支付方后重新预检。"
                        });
                    }
                }

                foreach (var group in groups)
                {
                    var key = ChongqingStage2Keys.SummaryKey(group.Entity, group.Kind);
                    if (!mainKeys.Contains(key))
                    {
                        issues.Add(new ChongqingStage2CheckIssue
                        {
                            Code = Stage2PreflightIssueKinds.NewSummarySubject,
                            Disposition = Stage2PreflightDisposition.Review,
                            Severity = "复核",
                            Category = "新增汇总主体默认字段",
                            Kind = Stage2PreflightIssueKinds.NewSummarySubject,
                            SettlementKind = group.Kind,
                            Owner = group.Owner,
                            Entity = group.Entity,
                            TemplateFile = options.SummaryTemplatePath,
                            SheetName = mainSheet.Name,
                            CurrentValue = "不委托；收款人=" + group.Entity + "；发票票种=平台；扣税率=台账唯一值；合计税率=13%",
                            Message = "程序将为新增主体写入已确认的最小必要汇总字段。",
                            Suggestion = "生成后请检查完整收款人、是否委托、发票票种、税率、负责人和借支字段。"
                        });
                        AddPaymentPartyRequirement(
                            options,
                            AvailablePaymentParties(workbook, options.Month),
                            mainSheet.Name,
                            group.Kind,
                            group.Entity,
                            group.Owner,
                            "新增汇总主体支付方选择",
                            "新增汇总主体没有可继承的支付方。",
                            issues);
                    }

                    EnrichNewCustomerIssues(group, sourceRows, issues);
                }

                AddMissingPaymentPartySheetIssues(options, workbook, mainRows, issues);
            }
        }

        private static void AddDuplicateSummaryIssues(
            ChongqingStage2Options options,
            IEnumerable<ChongqingSummaryMetaRow> sourceRows,
            IList<ChongqingStage2CheckIssue> issues)
        {
            foreach (var duplicate in sourceRows
                .GroupBy(item => item.SheetName + "|" + ChongqingStage2Keys.SummaryKey(item.Entity, item.Kind))
                .Where(sourceGroup => sourceGroup.Count() > 1))
            {
                var first = duplicate.First();
                issues.Add(new ChongqingStage2CheckIssue
                {
                    Code = Stage2PreflightIssueKinds.DuplicateSummarySubject,
                    Disposition = Stage2PreflightDisposition.Blocker,
                    Severity = "阻断",
                    Category = "汇总主体重复",
                    Kind = Stage2PreflightIssueKinds.DuplicateSummarySubject,
                    SettlementKind = first.Kind,
                    Entity = first.Entity,
                    TemplateFile = options.SummaryTemplatePath,
                    SheetName = first.SheetName,
                    CurrentValue = string.Join("、", duplicate.Select(item => "第" + item.Row + "行")),
                    Message = first.SheetName + "中同一费用类型和主体出现多条记录，无法可靠继承长期字段。",
                    Suggestion = "请合并或删除重复主体行后重新预检。"
                });
            }
        }

        private static void AddPayeeSourceIssue(
            ChongqingStage2Options options,
            IList<ChongqingSummaryMetaRow> sourceRows,
            IList<ChongqingStage2CheckIssue> issues)
        {
            var first = sourceRows[0];
            var distinctPayees = sourceRows
                .Where(item => !string.IsNullOrEmpty(Stage2OpaqueText.NormalizeForComparison(item.Payee)))
                .GroupBy(item => Stage2OpaqueText.NormalizeForComparison(item.Payee), StringComparer.Ordinal)
                .Select(group => group.First().Payee)
                .ToList();
            if (distinctPayees.Count > 1)
            {
                issues.Add(new ChongqingStage2CheckIssue
                {
                    Code = Stage2PreflightIssueKinds.ConflictingPayees,
                    Disposition = Stage2PreflightDisposition.Blocker,
                    Severity = "阻断",
                    Category = "完整收款人字段冲突",
                    Kind = Stage2PreflightIssueKinds.ConflictingPayees,
                    SettlementKind = first.Kind,
                    Entity = first.Entity,
                    TemplateFile = options.SummaryTemplatePath,
                    SheetName = string.Join("、", sourceRows.Select(item => item.SheetName).Distinct()),
                    PreviousValue = string.Join("；", distinctPayees),
                    Message = "同一汇总主体从可靠来源读到多个互相冲突的完整收款人文本。",
                    Suggestion = "请统一完整收款人单元格；程序不会拆分、重排或解析其中姓名。"
                });
                return;
            }

            if (!sourceRows.Any(item => string.IsNullOrEmpty(Stage2OpaqueText.NormalizeForComparison(item.Payee))))
            {
                return;
            }

            issues.Add(new ChongqingStage2CheckIssue
            {
                Code = Stage2PreflightIssueKinds.PayeeSourceMissing,
                Disposition = Stage2PreflightDisposition.Review,
                Severity = "复核",
                Category = "收款人可靠来源有缺失",
                Kind = Stage2PreflightIssueKinds.PayeeSourceMissing,
                SettlementKind = first.Kind,
                Entity = first.Entity,
                TemplateFile = options.SummaryTemplatePath,
                SheetName = string.Join("、", sourceRows.Select(item => item.SheetName).Distinct()),
                PreviousValue = distinctPayees.Count == 0 ? "所有可靠来源均为空" : "存在空白来源",
                CurrentValue = distinctPayees.Count == 0 ? string.Empty : "将沿用唯一非空完整文本：" + distinctPayees[0],
                Message = distinctPayees.Count == 0
                    ? "存量主体的完整收款人在所有可靠来源中均为空，程序保持空白但允许继续生成。"
                    : "部分可靠来源收款人为空，程序将原样沿用唯一非空完整文本。",
                Suggestion = "生成后请核对完整收款人单元格和委托收款资料。"
            });
        }

        private static void AddPaymentPartyConflictIssue(
            ChongqingStage2Options options,
            IList<ChongqingSummaryMetaRow> sourceRows,
            IList<ChongqingStage2CheckIssue> issues)
        {
            var paymentParties = sourceRows
                .SelectMany(item => new[] { TextUtil.S(item.PaymentParty), TextUtil.S(item.SourcePaymentParty) })
                .Where(item => !string.IsNullOrWhiteSpace(item))
                .Distinct(StringComparer.Ordinal)
                .ToList();
            if (paymentParties.Count <= 1)
            {
                return;
            }

            var first = sourceRows[0];
            issues.Add(new ChongqingStage2CheckIssue
            {
                Code = Stage2PreflightIssueKinds.ConflictingPaymentParties,
                Disposition = Stage2PreflightDisposition.Blocker,
                Severity = "阻断",
                Category = "支付方来源冲突",
                Kind = Stage2PreflightIssueKinds.ConflictingPaymentParties,
                SettlementKind = first.Kind,
                Entity = first.Entity,
                TemplateFile = options.SummaryTemplatePath,
                SheetName = string.Join("、", sourceRows.Select(item => item.SheetName).Distinct()),
                PreviousValue = string.Join("、", paymentParties),
                Message = "同一汇总主体在主表字段或支付方工作表归属中出现多个支付方。",
                Suggestion = "请先统一主汇总表和支付方工作表归属后重新预检。"
            });
        }

        private static void AddPaymentPartyRequirement(
            ChongqingStage2Options options,
            IList<string> availablePaymentParties,
            string sheetName,
            string kind,
            string entity,
            string owner,
            string category,
            string message,
            IList<ChongqingStage2CheckIssue> issues)
        {
            var key = ChongqingStage2Keys.SummaryKey(entity, kind);
            if (issues.Any(existingIssue =>
                existingIssue.Code == Stage2PreflightIssueKinds.PaymentPartyRequired
                && ChongqingStage2Keys.SummaryKey(existingIssue.Entity, existingIssue.SettlementKind) == key))
            {
                return;
            }

            if (availablePaymentParties == null || availablePaymentParties.Count == 0)
            {
                issues.Add(new ChongqingStage2CheckIssue
                {
                    Code = Stage2PreflightIssueKinds.SummaryPaymentSheetMissing,
                    Disposition = Stage2PreflightDisposition.Blocker,
                    Severity = "阻断",
                    Category = "支付方月度工作表没有可靠来源",
                    Kind = Stage2PreflightIssueKinds.SummaryPaymentSheetMissing,
                    SettlementKind = kind,
                    Owner = owner,
                    Entity = entity,
                    TemplateFile = options.SummaryTemplatePath,
                    SheetName = sheetName,
                    Message = message + " 但清能和清辉均没有目标月工作表或可克隆的较早月份来源。",
                    Suggestion = "请先补充至少一个可靠支付方月度工作表来源后重新预检。"
                });
                return;
            }

            var issue = new ChongqingStage2CheckIssue
            {
                Code = Stage2PreflightIssueKinds.PaymentPartyRequired,
                Disposition = Stage2PreflightDisposition.RequiredDecision,
                Severity = "确认",
                Category = category,
                Kind = ChongqingStage2IssueKinds.NewSummarySubjectPaymentPartyRequired,
                SettlementKind = kind,
                Owner = owner,
                Entity = entity,
                TemplateFile = options.SummaryTemplatePath,
                SheetName = sheetName,
                CurrentValue = "可选支付方：" + string.Join("、", availablePaymentParties),
                Message = message,
                Suggestion = availablePaymentParties.Count == ChongqingStage2PaymentParties.Supported.Count()
                    ? "请选择清能或清辉；本次选择只写入当前输出汇总表副本。"
                    : "仅列出模板中有可靠月度来源的支付方；若业务应选其它支付方，请先补充对应模板后重新预检。",
                RequiresPaymentPartySelection = true
            };
            issue.AvailablePaymentParties.AddRange(availablePaymentParties);
            issues.Add(issue);
        }

        private static List<string> AvailablePaymentParties(XLWorkbook workbook, int month)
        {
            return ChongqingStage2PaymentParties.Supported
                .Where(paymentParty => workbook.Worksheets.Any(sheet =>
                        sheet.Name == paymentParty + month + "月")
                    || PreviousPaymentPartySheet(workbook, paymentParty, month) != null)
                .ToList();
        }

        private static void EnrichNewCustomerIssues(
            GroupSettlementTotal group,
            IList<ChongqingSummaryMetaRow> sourceRows,
            IList<ChongqingStage2CheckIssue> issues)
        {
            var key = ChongqingStage2Keys.SummaryKey(group.Entity, group.Kind);
            var matches = sourceRows
                .Where(item => ChongqingStage2Keys.SummaryKey(item.Entity, item.Kind) == key)
                .ToList();
            var inheritedPayee = matches
                .Select(item => item.Payee)
                .FirstOrDefault(item => !string.IsNullOrEmpty(Stage2OpaqueText.NormalizeForComparison(item)));
            var historicalParties = matches
                .SelectMany(item => new[] { TextUtil.S(item.PaymentParty), TextUtil.S(item.SourcePaymentParty) })
                .Where(ChongqingStage2PaymentParties.Supported.Contains)
                .Distinct(StringComparer.Ordinal)
                .ToList();
            var inheritedParty = historicalParties.Count == 1 ? historicalParties[0] : null;

            foreach (var issue in issues.Where(issue =>
                issue.Code == Stage2PreflightIssueKinds.NewCustomer
                && ChongqingStage2Keys.SummaryKey(issue.Entity, issue.SettlementKind) == key))
            {
                issue.CurrentValue = "继承收款人："
                    + (string.IsNullOrWhiteSpace(inheritedPayee) ? "空白" : inheritedPayee)
                    + "；支付方："
                    + (string.IsNullOrWhiteSpace(inheritedParty) ? "待选择" : inheritedParty);
                issue.Suggestion = string.IsNullOrWhiteSpace(inheritedParty)
                    ? "这是该主体下的新增客户；支付方没有唯一可靠历史值，需在本次预检中选择。请同时检查完整收款人和客户关系参数。"
                    : "这是存量主体下的新增客户；将继承该主体完整收款人和唯一可靠历史支付方。请同时检查客户关系参数。";
            }
        }

        private static void AddMissingPaymentPartySheetIssues(
            ChongqingStage2Options options,
            XLWorkbook workbook,
            IList<ChongqingSummaryMetaRow> mainRows,
            IList<ChongqingStage2CheckIssue> issues)
        {
            var requiredParties = new HashSet<string>(mainRows
                .Select(item => TextUtil.S(item.PaymentParty))
                .Where(ChongqingStage2PaymentParties.Supported.Contains));

            foreach (var paymentParty in requiredParties)
            {
                var target = workbook.Worksheets.FirstOrDefault(sheet =>
                    sheet.Name == paymentParty + options.Month + "月");
                var source = target ?? PreviousPaymentPartySheet(workbook, paymentParty, options.Month);
                if (source != null)
                {
                    continue;
                }

                issues.Add(new ChongqingStage2CheckIssue
                {
                    Code = Stage2PreflightIssueKinds.SummaryPaymentSheetMissing,
                    Disposition = Stage2PreflightDisposition.Blocker,
                    Severity = "阻断",
                    Category = "支付方月度工作表缺少可靠来源",
                    Kind = Stage2PreflightIssueKinds.SummaryPaymentSheetMissing,
                    TemplateFile = options.SummaryTemplatePath,
                    SheetName = paymentParty + options.Month + "月",
                    CurrentValue = paymentParty,
                    Message = "本月存在归属“" + paymentParty + "”的汇总主体，但模板既没有目标月份工作表，也没有可克隆的较早月份工作表。",
                    Suggestion = "请补充可靠的历史“" + paymentParty + "X月”工作表后重新预检。"
                });
            }
        }

        private static void AddSummaryTargetMonthIssues(
            ChongqingStage2Options options,
            XLWorkbook workbook,
            IXLWorksheet mainSheet,
            IList<ChongqingStage2CheckIssue> issues)
        {
            AddSummaryTargetMonthBlockIssue(options, mainSheet, mainSheet.Name, issues);

            foreach (var paymentParty in ChongqingStage2PaymentParties.Supported)
            {
                var targetName = paymentParty + options.Month + "月";
                var worksheet = workbook.Worksheets.FirstOrDefault(sheet => sheet.Name == targetName)
                    ?? PreviousPaymentPartySheet(workbook, paymentParty, options.Month);
                if (worksheet != null)
                {
                    AddSummaryTargetMonthBlockIssue(options, worksheet, targetName, issues);
                }
            }
        }

        private static void AddSummaryTargetMonthBlockIssue(
            ChongqingStage2Options options,
            IXLWorksheet worksheet,
            string outputSheetName,
            IList<ChongqingStage2CheckIssue> issues)
        {
            var monthColumns = FindSummaryMonthBlocks(worksheet, options.Month);
            if (monthColumns.Count == 0)
            {
                return;
            }

            var ambiguous = monthColumns.Count > 1;
            issues.Add(new ChongqingStage2CheckIssue
            {
                Code = Stage2PreflightIssueKinds.SummaryTargetMonthAlreadyExists,
                Disposition = ambiguous
                    ? Stage2PreflightDisposition.Blocker
                    : Stage2PreflightDisposition.Review,
                Severity = ambiguous ? "阻断" : "复核",
                Category = ambiguous ? "汇总表目标月份块重复" : "汇总表目标月份块已存在",
                Kind = Stage2PreflightIssueKinds.SummaryTargetMonthAlreadyExists,
                TemplateFile = options.SummaryTemplatePath,
                SheetName = worksheet.Name,
                CurrentValue = string.Join("、", monthColumns.Select(column =>
                    "第" + column.ToString(CultureInfo.InvariantCulture) + "列")),
                Message = ambiguous
                    ? "本次将用于生成“" + outputSheetName + "”的工作表“" + worksheet.Name
                        + "”包含多个目标月份块，程序无法确定应重写哪一个。"
                    : "本次将用于生成“" + outputSheetName + "”的工作表“" + worksheet.Name
                        + "”已经包含唯一目标月份块，程序将保留结构并安全重写本月业务区。",
                Suggestion = ambiguous
                    ? "请只保留一个正确的目标月份块后重新预检。"
                    : "请确认该目标月占位或人工内容可以被本次结果覆盖。"
            });
        }

        private static void WriteSummarySheet(
            IXLWorksheet worksheet,
            IList<GroupSettlementTotal> groups,
            int month,
            ISet<string> allowedKeys,
            IList<string> warnings,
            IDictionary<string, string> paymentPartyByKey,
            IDictionary<string, double> taxRateByKey,
            IDictionary<string, string> payeeByKey,
            ISet<string> newSubjectKeys,
            IXLWorksheet canonicalSheet,
            bool refreshPaymentPartyVisibility = false)
        {
            DeleteSummaryRowsNotAllowed(worksheet, allowedKeys);
            var monthColumn = FindOrInsertSummaryMonthBlock(worksheet, month);
            var cumulativeColumn = SummaryColumn(worksheet, "当年费用总计");
            if (allowedKeys != null && refreshPaymentPartyVisibility)
            {
                ApplyPaymentPartySheetVisibility(worksheet, monthColumn, cumulativeColumn);
            }

            if (canonicalSheet != null)
            {
                EnsureCanonicalRowsPresent(worksheet, canonicalSheet, allowedKeys);
            }

            var totalByKey = groups.ToDictionary(group => ChongqingStage2Keys.SummaryKey(group.Entity, group.Kind), group => group);
            var knownKeys = new HashSet<string>(ReadSummaryMeta(worksheet).Select(row => ChongqingStage2Keys.SummaryKey(row.Entity, row.Kind)));
            var newGroups = groups.Where(group => !knownKeys.Contains(ChongqingStage2Keys.SummaryKey(group.Entity, group.Kind))).ToList();
            InsertNewSummaryRows(
                worksheet,
                newGroups,
                warnings,
                paymentPartyByKey,
                taxRateByKey,
                payeeByKey,
                canonicalSheet);

            var rows = ReadSummaryMeta(worksheet).OrderBy(row => row.Row).ToList();
            for (var index = 0; index < rows.Count; index++)
            {
                var row = rows[index].Row;
                var info = rows[index];
                GroupSettlementTotal total;
                totalByKey.TryGetValue(ChongqingStage2Keys.SummaryKey(info.Entity, info.Kind), out total);
                worksheet.Cell(row, 1).Value = index + 1;
                if (canonicalSheet != null)
                {
                    SyncCanonicalLongTermFields(
                        canonicalSheet,
                        FindUniqueSummaryRow(canonicalSheet, info.Entity, info.Kind),
                        worksheet,
                        row);
                }

                WriteSummaryValues(
                    worksheet,
                    row,
                    monthColumn,
                    cumulativeColumn,
                    total,
                    info.Entity,
                    info.Kind,
                    month,
                    newSubjectKeys.Contains(ChongqingStage2Keys.SummaryKey(info.Entity, info.Kind)),
                    paymentPartyByKey,
                    taxRateByKey,
                    payeeByKey);
            }

            var totalRow = FindSummaryTotalRow(worksheet);
            worksheet.Cell(totalRow, 1).Value = "合计";
            WriteSummaryTotalRow(worksheet, totalRow, cumulativeColumn);
        }

        private static void EnsureCanonicalRowsPresent(
            IXLWorksheet worksheet,
            IXLWorksheet canonicalSheet,
            ISet<string> allowedKeys)
        {
            if (canonicalSheet == null || allowedKeys == null)
            {
                return;
            }

            var existingKeys = new HashSet<string>(ReadSummaryMeta(worksheet).Select(item =>
                ChongqingStage2Keys.SummaryKey(item.Entity, item.Kind)));
            var totalRow = FindSummaryTotalRow(worksheet);
            foreach (var canonical in ReadSummaryMeta(canonicalSheet)
                .Where(item =>
                    allowedKeys.Contains(ChongqingStage2Keys.SummaryKey(item.Entity, item.Kind))
                    && !existingKeys.Contains(ChongqingStage2Keys.SummaryKey(item.Entity, item.Kind)))
                .OrderBy(item => item.Row))
            {
                worksheet.Row(totalRow).InsertRowsAbove(1);
                canonicalSheet.Row(canonical.Row).CopyTo(worksheet.Row(totalRow));
                existingKeys.Add(ChongqingStage2Keys.SummaryKey(canonical.Entity, canonical.Kind));
                totalRow++;
            }
        }

        private static void SyncCanonicalLongTermFields(
            IXLWorksheet canonicalSheet,
            int canonicalRow,
            IXLWorksheet targetSheet,
            int targetRow)
        {
            if (canonicalRow <= 0)
            {
                throw new InvalidDataException("重庆支付方工作表存在主汇总表无法唯一定位的主体行。");
            }

            for (var column = 4; column <= 11; column++)
            {
                canonicalSheet.Cell(canonicalRow, column).CopyTo(targetSheet.Cell(targetRow, column));
            }

            foreach (var header in new[]
            {
                "当年费用总计",
                "借支",
                "已抵扣借支",
                "借支剩余未抵扣",
                "借支开始抵扣月份",
                "借支还完月份",
                "代理/居间/退补电费新增月份",
                "支付方",
                "备注"
            })
            {
                var sourceColumn = SummaryColumn(canonicalSheet, header);
                var targetColumn = SummaryColumn(targetSheet, header);
                canonicalSheet.Cell(canonicalRow, sourceColumn)
                    .CopyTo(targetSheet.Cell(targetRow, targetColumn));
            }
        }

        private static void VerifyPaymentPartyMembership(
            XLWorkbook workbook,
            int month,
            GroupSettlementTotal group,
            string expectedPaymentParty,
            string expectedPayee,
            double expectedTaxRate)
        {
            var key = ChongqingStage2Keys.SummaryKey(group.Entity, group.Kind);
            foreach (var paymentParty in ChongqingStage2PaymentParties.Supported)
            {
                var sheetName = paymentParty + month + "月";
                var worksheet = workbook.Worksheets.FirstOrDefault(sheet => sheet.Name == sheetName);
                if (worksheet == null)
                {
                    throw new InvalidDataException("重庆阶段二汇总表缺少支付方工作表：" + sheetName);
                }

                var monthColumns = FindSummaryMonthBlocks(worksheet, month);
                if (monthColumns.Count != 1)
                {
                    throw new InvalidDataException(
                        "重庆阶段二支付方工作表的目标月份块数量不是 1："
                        + sheetName + "，实际=" + monthColumns.Count + "。");
                }

                var matches = ReadSummaryMeta(worksheet)
                    .Where(row => ChongqingStage2Keys.SummaryKey(row.Entity, row.Kind) == key)
                    .ToList();
                var expectedCount = paymentParty == expectedPaymentParty ? 1 : 0;
                if (matches.Count != expectedCount)
                {
                    throw new InvalidDataException(
                        "重庆阶段二支付方工作表归属不互斥："
                        + group.Kind + " " + group.Entity + "，" + sheetName);
                }

                if (matches.Count == 1
                    && !Stage2OpaqueText.AreEquivalent(matches[0].Payee, expectedPayee))
                {
                    throw new InvalidDataException(
                        "重庆阶段二支付方工作表的完整收款人未与主汇总保持一致："
                        + group.Kind + " " + group.Entity + "，" + sheetName);
                }

                if (matches.Count != 1)
                {
                    continue;
                }

                var match = matches[0];
                var monthColumn = monthColumns[0];

                var amount = ClosedXmlUtil.CellNumber(
                    worksheet.Cell(match.Row, monthColumn + SummaryKindOffset(group.Kind)));
                var withholdingTaxRate = ClosedXmlUtil.CellNumber(worksheet.Cell(match.Row, 9));
                var totalTaxRate = ClosedXmlUtil.CellNumber(worksheet.Cell(match.Row, 10));
                var expectedFormula = "J" + match.Row + "-I" + match.Row;
                var actualOwner = ChongqingStage2ExcelUtil.CellText(worksheet.Cell(match.Row, 11));
                if (Math.Abs(amount - group.ExpectedNet) > Stage2SettlementCalculator.AmountTolerance
                    || !ChongqingStage2ExcelUtil.TaxRatesEqual(withholdingTaxRate, expectedTaxRate)
                    || !ChongqingStage2ExcelUtil.TaxRatesEqual(totalTaxRate, 0.13d)
                    || !string.Equals(TextUtil.S(worksheet.Cell(match.Row, 8).FormulaA1), expectedFormula, StringComparison.OrdinalIgnoreCase)
                    || TextUtil.CustomerKey(actualOwner) != TextUtil.CustomerKey(group.Owner)
                    || TextUtil.S(match.PaymentParty) != expectedPaymentParty)
                {
                    throw new InvalidDataException(
                        "重庆阶段二支付方工作表金额或长期关键字段未与主汇总保持一致："
                        + group.Kind + " " + group.Entity + "，" + sheetName);
                }
            }
        }

        private static List<ChongqingSummaryMetaRow> ReadReliableSummaryMeta(
            XLWorkbook workbook,
            IXLWorksheet mainSheet,
            int month)
        {
            var result = new List<ChongqingSummaryMetaRow>();
            result.AddRange(ReadSummaryMeta(mainSheet));
            foreach (var paymentParty in ChongqingStage2PaymentParties.Supported)
            {
                var targetName = paymentParty + month + "月";
                var sheet = workbook.Worksheets.FirstOrDefault(item => item.Name == targetName)
                    ?? PreviousPaymentPartySheet(workbook, paymentParty, month);
                if (sheet != null && sheet.Name != mainSheet.Name)
                {
                    var rows = ReadSummaryMeta(sheet);
                    foreach (var row in rows)
                    {
                        row.SourcePaymentParty = paymentParty;
                    }

                    result.AddRange(rows);
                }
            }

            return result;
        }

        private static void WriteSummaryValues(
            IXLWorksheet worksheet,
            int row,
            int monthColumn,
            int cumulativeColumn,
            GroupSettlementTotal total,
            string entity,
            string kind,
            int month,
            bool isNewSubject,
            IDictionary<string, string> paymentPartyByKey,
            IDictionary<string, double> taxRateByKey,
            IDictionary<string, string> payeeByKey)
        {
            worksheet.Cell(row, monthColumn).Clear(XLClearOptions.Contents);
            worksheet.Cell(row, monthColumn + 1).Clear(XLClearOptions.Contents);
            worksheet.Cell(row, monthColumn + 2).Clear(XLClearOptions.Contents);

            var key = ChongqingStage2Keys.SummaryKey(entity, kind);
            if (paymentPartyByKey.ContainsKey(key))
            {
                worksheet.Cell(row, cumulativeColumn + 7).Value = paymentPartyByKey[key];
            }

            if (payeeByKey.ContainsKey(key))
            {
                worksheet.Cell(row, 6).Value = payeeByKey[key] ?? string.Empty;
            }

            if (total != null)
            {
                worksheet.Cell(row, monthColumn + SummaryKindOffset(total.Kind)).Value = Math.Round(total.ExpectedNet, 4);
                WriteTaxValues(worksheet, row, TaxRateFromIndex(taxRateByKey, entity, kind));
                worksheet.Cell(row, 11).Value = total.Owner;
            }

            worksheet.Cell(row, monthColumn + 3).FormulaA1 = SumFormula(row, monthColumn, monthColumn + 2);
            worksheet.Cell(row, monthColumn + 5).FormulaA1 = ClosedXmlUtil.ColumnLetter(monthColumn + 3) + row + "-" + ClosedXmlUtil.ColumnLetter(monthColumn + 4) + row;

            if (isNewSubject)
            {
                var firstMonthColumn = FirstSummaryMonthColumn(worksheet);
                worksheet.Cell(row, cumulativeColumn).FormulaA1 = SumEverySix(row, firstMonthColumn, cumulativeColumn - 1, 3);
                worksheet.Cell(row, cumulativeColumn + 2).FormulaA1 = SumEverySix(row, firstMonthColumn, cumulativeColumn - 1, 4);
                worksheet.Cell(row, cumulativeColumn + 3).FormulaA1 = ClosedXmlUtil.ColumnLetter(cumulativeColumn + 1) + row + "-" + ClosedXmlUtil.ColumnLetter(cumulativeColumn + 2) + row;
                worksheet.Cell(row, cumulativeColumn + 6).Value = new DateTime(ChongqingStage2Layout.Year, month, 1);
            }
        }

        private static void InsertNewSummaryRows(
            IXLWorksheet worksheet,
            IList<GroupSettlementTotal> newGroups,
            IList<string> warnings,
            IDictionary<string, string> paymentPartyByKey,
            IDictionary<string, double> taxRateByKey,
            IDictionary<string, string> payeeByKey,
            IXLWorksheet canonicalSheet)
        {
            if (newGroups.Count == 0)
            {
                return;
            }

            var totalRow = FindSummaryTotalRow(worksheet);
            var templateRow = Math.Max(ChongqingStage2Layout.SummaryDataStartRow, totalRow - 1);
            foreach (var group in newGroups)
            {
                worksheet.Row(totalRow).InsertRowsAbove(1);
                var canonicalRow = FindUniqueSummaryRow(canonicalSheet, group.Entity, group.Kind);
                if (canonicalRow > 0)
                {
                    canonicalSheet.Row(canonicalRow).CopyTo(worksheet.Row(totalRow));
                }
                else
                {
                    worksheet.Row(templateRow).CopyTo(worksheet.Row(totalRow));
                    worksheet.Row(totalRow).Clear(XLClearOptions.Contents);
                }

                worksheet.Cell(totalRow, 1).Value = totalRow - ChongqingStage2Layout.SummaryDataStartRow + 1;
                worksheet.Cell(totalRow, 2).Value = group.Entity;
                worksheet.Cell(totalRow, 3).Value = group.Kind;
                if (canonicalRow <= 0)
                {
                    worksheet.Cell(totalRow, 4).Value = "否";
                    worksheet.Cell(totalRow, 7).Value = "平台";
                }

                worksheet.Cell(totalRow, 6).Value = PayeeFromIndex(payeeByKey, group.Entity, group.Kind);
                WriteTaxValues(
                    worksheet,
                    totalRow,
                    TaxRateFromIndex(taxRateByKey, group.Entity, group.Kind));
                worksheet.Cell(totalRow, 11).Value = group.Owner;
                var cumulativeColumn = SummaryColumn(worksheet, "当年费用总计");
                worksheet.Cell(totalRow, cumulativeColumn + 7).Value = PaymentPartyFromIndex(paymentPartyByKey, group.Entity, group.Kind);
                if (canonicalSheet == null)
                {
                    warnings.Add("新增重庆汇总主体：" + group.Kind + " " + group.Entity + "（负责人：" + group.Owner + "；支付方：" + PaymentPartyFromIndex(paymentPartyByKey, group.Entity, group.Kind) + "），已写入“不委托、收款人=主体、平台票种、台账扣税率、13%合计税率”等默认字段，请人工复核。");
                }
                totalRow++;
            }
        }

        private static int FindUniqueSummaryRow(
            IXLWorksheet worksheet,
            string entity,
            string kind)
        {
            if (worksheet == null)
            {
                return 0;
            }

            var key = ChongqingStage2Keys.SummaryKey(entity, kind);
            var matches = ReadSummaryMeta(worksheet)
                .Where(item => ChongqingStage2Keys.SummaryKey(item.Entity, item.Kind) == key)
                .ToList();
            return matches.Count == 1 ? matches[0].Row : 0;
        }

        private static IDictionary<string, double> BuildTaxRateIndex(
            IList<ChongqingSettlementDetail> details)
        {
            var result = new Dictionary<string, double>();
            foreach (var group in details.GroupBy(detail =>
                ChongqingStage2Keys.SummaryKey(detail.Entity, detail.Kind)))
            {
                var rates = group
                    .Select(detail => detail.TaxRate)
                    .Distinct(new TaxRateComparer())
                    .ToList();
                if (rates.Count != 1)
                {
                    var first = group.First();
                    throw new InvalidDataException("重庆阶段二无法确定唯一扣税率：" + first.Kind + " " + first.Entity);
                }

                result[group.Key] = rates[0];
            }

            return result;
        }

        private static IDictionary<string, string> BuildPayeeIndex(
            IList<GroupSettlementTotal> groups,
            IList<ChongqingSummaryMetaRow> mainRows,
            IList<ChongqingSummaryMetaRow> sourceRows)
        {
            var result = new Dictionary<string, string>();
            var identities = mainRows
                .Select(item => new
                {
                    Key = ChongqingStage2Keys.SummaryKey(item.Entity, item.Kind),
                    item.Entity,
                    item.Kind
                })
                .Concat(groups.Select(item => new
                {
                    Key = ChongqingStage2Keys.SummaryKey(item.Entity, item.Kind),
                    item.Entity,
                    item.Kind
                }))
                .GroupBy(item => item.Key)
                .Select(group => group.First())
                .ToList();
            foreach (var identity in identities)
            {
                var key = identity.Key;
                var matchingMainRows = mainRows
                    .Where(item => ChongqingStage2Keys.SummaryKey(item.Entity, item.Kind) == key)
                    .ToList();
                var matchingSources = sourceRows
                    .Where(item => ChongqingStage2Keys.SummaryKey(item.Entity, item.Kind) == key)
                    .ToList();
                var distinctPayees = matchingSources
                    .Where(item => !string.IsNullOrEmpty(Stage2OpaqueText.NormalizeForComparison(item.Payee)))
                    .GroupBy(item => Stage2OpaqueText.NormalizeForComparison(item.Payee), StringComparer.Ordinal)
                    .Select(payeeGroup => payeeGroup.First().Payee)
                    .ToList();
                if (distinctPayees.Count > 1)
                {
                    throw new InvalidDataException("重庆阶段二无法确定唯一完整收款人文本：" + identity.Kind + " " + identity.Entity);
                }

                var mainPayee = matchingMainRows
                    .Select(item => item.Payee)
                    .FirstOrDefault(item => !string.IsNullOrEmpty(Stage2OpaqueText.NormalizeForComparison(item)));
                if (mainPayee != null)
                {
                    result[key] = mainPayee;
                }
                else if (distinctPayees.Count == 1)
                {
                    result[key] = distinctPayees[0];
                }
                else if (matchingMainRows.Count > 0)
                {
                    result[key] = matchingMainRows[0].Payee ?? string.Empty;
                }
                else
                {
                    result[key] = identity.Entity;
                }
            }

            return result;
        }

        private static void WriteTaxValues(IXLWorksheet worksheet, int row, double withholdingTaxRate)
        {
            worksheet.Cell(row, 8).FormulaA1 = "J" + row + "-I" + row;
            worksheet.Cell(row, 9).Value = Math.Round(withholdingTaxRate, 10);
            worksheet.Cell(row, 10).Value = 0.13d;
        }

        private static double TaxRateFromIndex(
            IDictionary<string, double> taxRateByKey,
            string entity,
            string kind)
        {
            double taxRate;
            if (taxRateByKey.TryGetValue(ChongqingStage2Keys.SummaryKey(entity, kind), out taxRate))
            {
                return taxRate;
            }

            throw new InvalidDataException("重庆阶段二汇总主体缺少扣税率：" + kind + " " + entity);
        }

        private static string PayeeFromIndex(
            IDictionary<string, string> payeeByKey,
            string entity,
            string kind)
        {
            string payee;
            if (payeeByKey.TryGetValue(ChongqingStage2Keys.SummaryKey(entity, kind), out payee))
            {
                return payee ?? string.Empty;
            }

            throw new InvalidDataException("重庆阶段二汇总主体缺少完整收款人值：" + kind + " " + entity);
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

        private static IDictionary<string, string> BuildPaymentPartyIndex(
            ChongqingStage2Options options,
            IList<GroupSettlementTotal> groups,
            IList<ChongqingSummaryMetaRow> mainMeta)
        {
            var partyByKey = new Dictionary<string, string>();
            foreach (var row in mainMeta)
            {
                var key = ChongqingStage2Keys.SummaryKey(row.Entity, row.Kind);
                if (ChongqingStage2PaymentParties.Supported.Contains(row.PaymentParty))
                {
                    partyByKey[key] = row.PaymentParty;
                }
            }

            foreach (var group in groups)
            {
                var key = ChongqingStage2Keys.SummaryKey(group.Entity, group.Kind);
                if (partyByKey.ContainsKey(key))
                {
                    continue;
                }

                string paymentParty;
                if (!TryGetPaymentPartyDecision(options, group.Entity, group.Kind, out paymentParty))
                {
                    throw new InvalidOperationException("重庆阶段二新增汇总主体支付方未选择：" + group.Kind + " " + group.Entity + "。请在预检窗口选择清能或清辉后再生成。");
                }

                partyByKey[key] = paymentParty;
            }

            return partyByKey;
        }

        private static IXLWorksheet PreparePaymentPartySheet(
            XLWorkbook workbook,
            string paymentParty,
            int month,
            IXLWorksheet mainSheet,
            out bool created)
        {
            var targetName = paymentParty + month + "月";
            var existing = workbook.Worksheets.FirstOrDefault(sheet => sheet.Name == targetName);
            if (existing != null)
            {
                created = false;
                return existing;
            }

            var source = PreviousPaymentPartySheet(workbook, paymentParty, month);
            if (source != null)
            {
                created = true;
                return source.CopyTo(targetName);
            }

            created = true;
            return mainSheet.CopyTo(targetName);
        }

        private static IXLWorksheet PreviousPaymentPartySheet(XLWorkbook workbook, string paymentParty, int month)
        {
            return workbook.Worksheets
                .Select(sheet =>
                {
                    int parsedMonth;
                    return new { Sheet = sheet, Matched = TryParsePaymentPartySheet(sheet.Name, paymentParty, out parsedMonth), Month = parsedMonth };
                })
                .Where(item => item.Matched && item.Month < month)
                .OrderBy(item => item.Month)
                .Select(item => item.Sheet)
                .LastOrDefault();
        }

        private static bool TryParsePaymentPartySheet(string name, string paymentParty, out int month)
        {
            month = 0;
            var match = Regex.Match(TextUtil.S(name), "^" + Regex.Escape(paymentParty) + "(\\d{1,2})月$");
            return match.Success && int.TryParse(match.Groups[1].Value, out month);
        }

        private static void MoveTargetPaymentPartySheetsAfterSummary(XLWorkbook workbook, int month, IXLWorksheet mainSheet)
        {
            var position = mainSheet.Position + 1;
            foreach (var paymentParty in ChongqingStage2PaymentParties.Supported)
            {
                var targetName = paymentParty + month + "月";
                var sheet = workbook.Worksheets.FirstOrDefault(item => item.Name == targetName);
                if (sheet == null)
                {
                    continue;
                }

                sheet.Position = position++;
            }
        }

        private static void ApplyPaymentPartyTitleMerge(IXLWorksheet worksheet)
        {
            var titleEndColumn = FindHeaderColumnInRows(worksheet, "备注", 1, 3);
            if (titleEndColumn <= 1)
            {
                return;
            }

            var titleRange = worksheet.Range(1, 1, 1, titleEndColumn);
            titleRange.Unmerge();
            titleRange.Merge();
            titleRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            titleRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            titleRange.Style.Font.Bold = true;
        }

        private static int FindOrInsertSummaryMonthBlock(IXLWorksheet worksheet, int month)
        {
            var existing = FindSummaryMonthBlocks(worksheet, month);
            if (existing.Count > 1)
            {
                throw new InvalidDataException(
                    worksheet.Name + " 中存在多个 " + month + " 月汇总区块，已停止生成。");
            }

            if (existing.Count == 1)
            {
                return existing[0];
            }

            var cumulativeColumn = SummaryColumn(worksheet, "当年费用总计");
            var insertAt = cumulativeColumn;
            worksheet.Column(insertAt).InsertColumnsBefore(6);
            for (var offset = 0; offset < 6; offset++)
            {
                var sourceColumn = Math.Max(1, insertAt - 6 + offset);
                worksheet.Column(sourceColumn).CopyTo(worksheet.Column(insertAt + offset));
                worksheet.Column(insertAt + offset).Unhide();
            }

            worksheet.Cell(2, insertAt).Value = new DateTime(ChongqingStage2Layout.Year, month, 1);
            worksheet.Cell(2, insertAt + 5).Value = "当月实际支付";
            worksheet.Cell(3, insertAt).Value = "代理费";
            worksheet.Cell(3, insertAt + 1).Value = "居间费";
            worksheet.Cell(3, insertAt + 2).Value = "退补电费";
            worksheet.Cell(3, insertAt + 3).Value = "费用合计";
            worksheet.Cell(3, insertAt + 4).Value = "当月抵扣";
            worksheet.Cell(3, insertAt + 5).Clear(XLClearOptions.Contents);
            return insertAt;
        }

        private static void ApplyPaymentPartySheetVisibility(IXLWorksheet worksheet, int monthColumn, int cumulativeColumn)
        {
            var lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? cumulativeColumn + 8;
            for (var column = 1; column <= lastColumn; column++)
            {
                if (ChongqingStage2ExcelUtil.CellText(worksheet.Cell(3, column)) != "代理费")
                {
                    continue;
                }

                for (var offset = 0; offset < 6 && column + offset <= lastColumn; offset++)
                {
                    worksheet.Column(column + offset).Hide();
                }
            }

            for (var offset = 0; offset < 6 && monthColumn + offset <= lastColumn; offset++)
            {
                worksheet.Column(monthColumn + offset).Unhide();
            }

            for (var column = cumulativeColumn; column <= Math.Min(lastColumn, cumulativeColumn + 6); column++)
            {
                worksheet.Column(column).Unhide();
            }

            if (cumulativeColumn + 7 <= lastColumn)
            {
                worksheet.Column(cumulativeColumn + 7).Hide();
            }

            if (cumulativeColumn + 8 <= lastColumn)
            {
                worksheet.Column(cumulativeColumn + 8).Unhide();
            }
        }

        internal static IList<int> FindSummaryMonthBlocks(IXLWorksheet worksheet, int month)
        {
            var result = new List<int>();
            var lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? 1;
            for (var column = 1; column <= lastColumn; column++)
            {
                if (ChongqingStage2ExcelUtil.CellText(worksheet.Cell(3, column)) != "代理费")
                {
                    continue;
                }

                if (SummaryHeaderMatchesMonth(worksheet.Cell(2, column), month))
                {
                    result.Add(column);
                }
            }

            return result;
        }

        private static bool SummaryHeaderMatchesMonth(IXLCell cell, int month)
        {
            try
            {
                if (cell.DataType == XLDataType.DateTime && cell.GetDateTime().Month == month && cell.GetDateTime().Year == ChongqingStage2Layout.Year)
                {
                    return true;
                }
            }
            catch
            {
            }

            var text = ChongqingStage2ExcelUtil.CellText(cell);
            return text.Contains(ChongqingStage2Layout.Year.ToString(CultureInfo.InvariantCulture))
                && (text.Contains(month + "月") || text.Contains("-" + month.ToString("00", CultureInfo.InvariantCulture) + "-"));
        }

        private static void DeleteSummaryRowsNotAllowed(IXLWorksheet worksheet, ISet<string> allowedKeys)
        {
            if (allowedKeys == null)
            {
                return;
            }

            var totalRow = FindSummaryTotalRow(worksheet);
            for (var row = totalRow - 1; row >= ChongqingStage2Layout.SummaryDataStartRow; row--)
            {
                var entity = ChongqingStage2ExcelUtil.CellText(worksheet.Cell(row, 2));
                var kind = ChongqingStage2ExcelUtil.CellText(worksheet.Cell(row, 3));
                if (string.IsNullOrWhiteSpace(entity) || !allowedKeys.Contains(ChongqingStage2Keys.SummaryKey(entity, kind)))
                {
                    worksheet.Row(row).Delete();
                }
            }
        }

        private static void WriteSummaryTotalRow(IXLWorksheet worksheet, int totalRow, int cumulativeColumn)
        {
            var firstMonthColumn = FirstSummaryMonthColumn(worksheet);
            for (var column = firstMonthColumn; column <= cumulativeColumn + 3; column++)
            {
                var letter = ClosedXmlUtil.ColumnLetter(column);
                worksheet.Cell(totalRow, column).FormulaA1 = ColumnSumFormula(letter, ChongqingStage2Layout.SummaryDataStartRow, totalRow - 1);
            }
        }

        private static string ColumnSumFormula(string letter, int firstRow, int lastRow)
        {
            return firstRow == lastRow
                ? "SUM(" + letter + firstRow + ")"
                : "SUM(" + letter + firstRow + ":" + letter + lastRow + ")";
        }

        private static int FirstSummaryMonthColumn(IXLWorksheet worksheet)
        {
            var lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? 1;
            for (var column = 1; column <= lastColumn; column++)
            {
                if (ChongqingStage2ExcelUtil.CellText(worksheet.Cell(3, column)) == "代理费")
                {
                    return column;
                }
            }

            throw new InvalidOperationException(worksheet.Name + " 未找到重庆汇总表月度费用列。");
        }

        private static string SumEverySix(int row, int firstColumn, int lastColumn, int offset)
        {
            var parts = new List<string>();
            for (var column = firstColumn + offset; column <= lastColumn; column += 6)
            {
                parts.Add(ClosedXmlUtil.ColumnLetter(column) + row);
            }

            return string.Join("+", parts);
        }

        private static string SumFormula(int row, int firstColumn, int lastColumn)
        {
            var parts = new List<string>();
            for (var column = firstColumn; column <= lastColumn; column++)
            {
                parts.Add(ClosedXmlUtil.ColumnLetter(column) + row);
            }

            return string.Join("+", parts);
        }

        private static int SummaryKindOffset(string kind)
        {
            if (kind == ChongqingStage2SettlementKinds.Proxy)
            {
                return 0;
            }

            if (kind == ChongqingStage2SettlementKinds.Intermediary)
            {
                return 1;
            }

            if (kind == ChongqingStage2SettlementKinds.Refund)
            {
                return 2;
            }

            throw new InvalidDataException("重庆阶段二汇总表包含不支持的费用类型：" + TextUtil.S(kind) + "。");
        }

        private static bool TryGetPaymentPartyDecision(ChongqingStage2Options options, string entity, string kind, out string paymentParty)
        {
            paymentParty = null;
            var key = ChongqingStage2Keys.SummaryKey(entity, kind);
            var decision = options.SummarySubjectDecisions
                .Where(item => item != null)
                .FirstOrDefault(item => ChongqingStage2Keys.SummaryKey(item.Entity, item.SettlementKind) == key);
            if (decision == null || string.IsNullOrWhiteSpace(decision.PaymentParty))
            {
                return false;
            }

            paymentParty = decision.PaymentParty;
            return true;
        }

        private static string PartyForSummaryTotal(GroupSettlementTotal total, IDictionary<string, string> partyByKey)
        {
            string party;
            if (partyByKey.TryGetValue(ChongqingStage2Keys.SummaryKey(total.Entity, total.Kind), out party))
            {
                return party;
            }

            throw new InvalidOperationException("重庆阶段二新增汇总主体支付方未选择：" + total.Kind + " " + total.Entity + "。");
        }

        private static string PaymentPartyFromIndex(IDictionary<string, string> paymentPartyByKey, string entity, string kind)
        {
            string paymentParty;
            if (paymentPartyByKey.TryGetValue(ChongqingStage2Keys.SummaryKey(entity, kind), out paymentParty))
            {
                return paymentParty;
            }

            throw new InvalidOperationException("重庆阶段二新增汇总主体支付方未选择：" + kind + " " + entity + "。");
        }

        private static IXLWorksheet FindSummaryWorksheet(XLWorkbook workbook)
        {
            var candidates = workbook.Worksheets
                .Where(sheet => !IsPaymentPartyMonthSheet(sheet.Name))
                .Where(IsReliableSummaryWorksheet)
                .ToList();
            if (candidates.Count == 1)
            {
                return candidates[0];
            }

            if (candidates.Count == 0)
            {
                throw new InvalidDataException("重庆阶段二汇总模板未找到唯一可靠的主汇总工作表；请确认存在“汇总表”，或仅有一张具备完整汇总表头的主表。");
            }

            throw new InvalidDataException(
                "重庆阶段二汇总模板找到多个可靠的主汇总工作表，无法自动选择："
                + string.Join("、", candidates.Select(sheet => sheet.Name)) + "。");
        }

        private static bool IsPaymentPartyMonthSheet(string sheetName)
        {
            foreach (var paymentParty in ChongqingStage2PaymentParties.Supported)
            {
                int month;
                if (TryParsePaymentPartySheet(sheetName, paymentParty, out month))
                {
                    return true;
                }
            }

            return false;
        }

        private static bool IsReliableSummaryWorksheet(IXLWorksheet worksheet)
        {
            return ChongqingStage2ExcelUtil.CellText(worksheet.Cell("A1")).Contains("汇总")
                && FindHeaderColumnInRows(worksheet, "名称", 1, 3) > 0
                && FindHeaderColumnInRows(worksheet, "类目", 1, 3) > 0
                && FindHeaderColumnInRows(worksheet, "当年费用总计", 1, 3) > 0;
        }

        private static bool IsSupportedSummaryKind(string kind)
        {
            return kind == ChongqingStage2SettlementKinds.Proxy
                || kind == ChongqingStage2SettlementKinds.Intermediary
                || kind == ChongqingStage2SettlementKinds.Refund;
        }

        private static List<ChongqingSummaryMetaRow> ReadSummaryMeta(IXLWorksheet worksheet)
        {
            var rows = new List<ChongqingSummaryMetaRow>();
            var totalRow = FindSummaryTotalRow(worksheet);
            var paymentPartyColumn = FindHeaderColumnInRows(worksheet, "支付方", 1, 3);
            for (var row = ChongqingStage2Layout.SummaryDataStartRow; row < totalRow; row++)
            {
                var entity = ChongqingStage2ExcelUtil.CellText(worksheet.Cell(row, 2));
                var kind = ChongqingStage2ExcelUtil.CellText(worksheet.Cell(row, 3));
                if (string.IsNullOrWhiteSpace(entity))
                {
                    continue;
                }

                if (!IsSupportedSummaryKind(kind))
                {
                    throw new InvalidDataException(
                        worksheet.Name + "第" + row + "行费用类型不受支持："
                        + (string.IsNullOrWhiteSpace(kind) ? "（空白）" : kind)
                        + "；仅允许代理费、居间费、退补电费。");
                }

                rows.Add(new ChongqingSummaryMetaRow
                {
                    Row = row,
                    Entity = entity,
                    Kind = kind,
                    Payee = worksheet.Cell(row, 6).GetFormattedString(),
                    PaymentParty = paymentPartyColumn > 0 ? ChongqingStage2ExcelUtil.CellText(worksheet.Cell(row, paymentPartyColumn)) : string.Empty,
                    SheetName = worksheet.Name
                });
            }

            return rows;
        }

        private static int FindSummaryTotalRow(IXLWorksheet worksheet)
        {
            var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? ChongqingStage2Layout.SummaryDataStartRow;
            for (var row = ChongqingStage2Layout.SummaryDataStartRow; row <= lastRow; row++)
            {
                if (ChongqingStage2ExcelUtil.CellText(worksheet.Cell(row, 1)) == "合计")
                {
                    return row;
                }
            }

            return lastRow + 1;
        }

        private static int SummaryColumn(IXLWorksheet worksheet, string header)
        {
            var column = FindHeaderColumnInRows(worksheet, header, 1, 3);
            if (column <= 0)
            {
                throw new InvalidOperationException(worksheet.Name + " 未找到重庆汇总表表头“" + header + "”。");
            }

            return column;
        }

        private static int FindHeaderColumnInRows(IXLWorksheet worksheet, string header, int firstRow, int lastRow)
        {
            var lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? 1;
            for (var row = firstRow; row <= lastRow; row++)
            {
                for (var column = 1; column <= lastColumn; column++)
                {
                    if (ChongqingStage2ExcelUtil.CellText(worksheet.Cell(row, column)) == header)
                    {
                        return column;
                    }
                }
            }

            return 0;
        }
    }
}
