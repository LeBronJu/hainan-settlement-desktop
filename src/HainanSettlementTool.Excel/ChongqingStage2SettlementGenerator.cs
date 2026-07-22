using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using HainanSettlementTool.Core.Models;
using HainanSettlementTool.Core.Services;

namespace HainanSettlementTool.Excel
{
    internal sealed class ChongqingStage2SettlementGenerator
    {
        public ChongqingStage2PreflightReport Analyze(ChongqingStage2Options options)
        {
            var baseInputBefore = CaptureInputFingerprint(options, null);
            var snapshot = ChongqingStage2LedgerReader.ReadSnapshot(options);
            var groups = ChongqingStage2LedgerReader.BuildGroups(snapshot.Details);
            var managedOutputPlan = ChongqingStage2SplitWorkbookWriter.BuildManagedOutputPlan(options, groups);
            EnsureInputFingerprintUnchanged(
                baseInputBefore,
                CaptureInputFingerprint(options, null),
                "预检读取期间");
            var observedOutputPaths = BuildObservedOutputPaths(options, managedOutputPlan);
            var inputBeforeIssues = CaptureInputFingerprint(options, observedOutputPaths);
            var report = new ChongqingStage2PreflightReport
            {
                Month = options.Month,
                SubjectCount = groups.Count
            };
            report.Issues.AddRange(BuildPreflightIssues(options, snapshot, groups, managedOutputPlan));
            var inputAfter = CaptureInputFingerprint(options, observedOutputPaths);
            EnsureInputFingerprintUnchanged(inputBeforeIssues, inputAfter, "预检期间");
            report.InputFingerprint = inputAfter;
            report.PreflightSignature = Stage2PreflightSignature.Create(
                options.Month,
                report.SubjectCount,
                report.Issues);
            return report;
        }

        public ChongqingStage2Report Generate(ChongqingStage2Options options)
        {
            EnsureGenerationWasAuthorizedByPreflight(options);
            var generationBaseFingerprint = CaptureInputFingerprint(options, null);
            var snapshot = ChongqingStage2LedgerReader.ReadSnapshot(options);
            var details = snapshot.Details;
            var groups = ChongqingStage2LedgerReader.BuildGroups(details);
            var managedOutputPlan = ChongqingStage2SplitWorkbookWriter.BuildManagedOutputPlan(options, groups);
            EnsureInputFingerprintUnchanged(
                generationBaseFingerprint,
                CaptureInputFingerprint(options, null),
                "生成读取期间");
            var observedOutputPaths = BuildObservedOutputPaths(options, managedOutputPlan);
            var generationInputFingerprint = CaptureInputFingerprint(options, observedOutputPaths);
            if (!Stage2InputFingerprint.Matches(
                options.ExpectedInputFingerprint,
                generationInputFingerprint))
            {
                throw new InvalidOperationException(
                    "重庆阶段二输入或计划覆盖的正式文件在预检确认后发生变化，本次未生成；请重新预检并确认。" );
            }

            var preflight = new ChongqingStage2PreflightReport
            {
                Month = options.Month,
                SubjectCount = groups.Count
            };
            preflight.Issues.AddRange(BuildPreflightIssues(options, snapshot, groups, managedOutputPlan));
            var preWorkspaceFingerprint = CaptureInputFingerprint(options, observedOutputPaths);
            EnsureInputFingerprintUnchanged(
                generationInputFingerprint,
                preWorkspaceFingerprint,
                "生成准备期间");
            preflight.InputFingerprint = preWorkspaceFingerprint;
            preflight.PreflightSignature = Stage2PreflightSignature.Create(
                options.Month,
                preflight.SubjectCount,
                preflight.Issues);
            if (!Stage2PreflightSignature.Matches(
                options.ExpectedPreflightSignature,
                preflight.PreflightSignature))
            {
                throw new InvalidOperationException(
                    "重庆阶段二预检项目在确认后发生变化，本次未生成；请重新预检并确认。" );
            }

            EnsurePreflightAllowsGeneration(options, preflight);

            var expectedKeys = new HashSet<string>(groups.Select(group =>
                ChongqingStage2Keys.SummaryKey(group.Entity, group.Kind)));
            var workspace = new Stage2BatchWorkspace(options.OutputDirectory, "重庆", options.Month);
            var stagingOptions = CreateStagingOptions(options, workspace.StagingDirectory);
            try
            {
                var warnings = new List<string>();
                var auditIssues = new List<ChongqingStage2CheckIssue>();
                AddPreflightAuditIssues(options, preflight, auditIssues);

                groups = ChongqingStage2SplitWorkbookWriter.BuildSplitFiles(
                    stagingOptions,
                    details,
                    warnings);
                EnsureGeneratedGroupsMatch(expectedKeys, groups);
                EnsureGeneratedOutputsMatchPlan(workspace, groups, managedOutputPlan);

                var stagedSummaryPath = ChongqingStage2SummaryWorkbookWriter.BuildSummary(
                    stagingOptions,
                    groups,
                    details,
                    warnings);
                ChongqingStage2SplitWorkbookWriter.VerifyGeneratedSplitWorkbooks(
                    groups,
                    details,
                    options.Month);
                ChongqingStage2SummaryWorkbookWriter.VerifyGeneratedSummary(
                    stagingOptions,
                    groups,
                    details,
                    stagedSummaryPath);

                var stagedSplitPaths = groups.Select(group => group.OutputFile).ToList();
                RemapOutputReferences(workspace, groups, warnings, auditIssues);
                var finalSummaryPath = workspace.GetFinalPath(stagedSummaryPath);
                var report = ChongqingStage2ReportWriter.CreateReport(
                    options,
                    details,
                    groups,
                    finalSummaryPath,
                    warnings,
                    auditIssues);
                var stagedReportPath = workspace.GetStagingPath(Path.GetFileName(report.ReportPath));
                var stagedValidationPath = workspace.GetStagingPath(Path.GetFileName(report.ValidationReportPath));
                ChongqingStage2ReportWriter.Write(report, stagedReportPath, stagedValidationPath);

                Stage2BatchIntegrityVerifier.VerifyFiles(
                    workspace,
                    stagedSplitPaths,
                    stagedSummaryPath,
                    new[] { stagedReportPath, stagedValidationPath });
                ChongqingStage2SplitWorkbookWriter.EnsureManagedOutputStillSafe(
                    options,
                    managedOutputPlan);
                EnsureInputFingerprintUnchanged(
                    generationInputFingerprint,
                    CaptureInputFingerprint(options, observedOutputPaths),
                    "生成期间");
                workspace.Publish(() =>
                {
                    ChongqingStage2SplitWorkbookWriter.EnsureManagedOutputStillSafe(
                        options,
                        managedOutputPlan);
                    EnsureInputFingerprintUnchanged(
                        generationInputFingerprint,
                        CaptureInputFingerprint(options, observedOutputPaths),
                        "发布锁内复核期间");
                });
                return report;
            }
            catch (Exception ex)
            {
                throw PreserveFailedWorkspace(workspace, ex, "重庆");
            }
        }

        private static List<ChongqingStage2CheckIssue> BuildPreflightIssues(
            ChongqingStage2Options options,
            ChongqingStage2LedgerSnapshot snapshot,
            IList<GroupSettlementTotal> groups,
            IList<ChongqingManagedOutputPlanItem> managedOutputPlan)
        {
            var issues = new List<ChongqingStage2CheckIssue>(snapshot.Issues);
            AddLedgerAmountDifferenceIssues(options, snapshot.Details, issues);
            ChongqingStage2SplitWorkbookWriter.AddTemplateIssues(options, groups, snapshot.Details, issues);
            ChongqingStage2SplitWorkbookWriter.AddManagedOutputIssues(options, managedOutputPlan, issues);
            ChongqingStage2SummaryWorkbookWriter.AddSummaryPaymentIssues(options, groups, issues);
            return issues;
        }

        private static void AddLedgerAmountDifferenceIssues(
            ChongqingStage2Options options,
            IEnumerable<ChongqingSettlementDetail> details,
            IList<ChongqingStage2CheckIssue> issues)
        {
            foreach (var detail in details.Where(item =>
                Math.Abs(item.LedgerNet - item.CalculatedNet) > Stage2SettlementCalculator.AmountTolerance))
            {
                issues.Add(new ChongqingStage2CheckIssue
                {
                    Code = Stage2PreflightIssueKinds.LedgerAmountDifference,
                    Disposition = Stage2PreflightDisposition.Review,
                    Severity = "复核",
                    Category = "台账与分表金额不一致",
                    Kind = Stage2PreflightIssueKinds.LedgerAmountDifference,
                    SettlementKind = detail.Kind,
                    Customer = detail.Customer,
                    Owner = detail.Owner,
                    Entity = detail.Entity,
                    LedgerRow = detail.LedgerRow,
                    TemplateFile = options.LedgerPath,
                    PreviousValue = "台账净额：" + Stage2SettlementCalculator.FormatAmount(detail.LedgerNet),
                    CurrentValue = "分表自算：" + Stage2SettlementCalculator.FormatAmount(detail.CalculatedNet),
                    Message = "重庆" + detail.Kind + "主体“" + detail.Entity + "”下客户“"
                        + detail.Customer + "”的台账净额与分表公式自算结果不一致。",
                    Suggestion = "请在生成前检查台账第" + detail.LedgerRow
                        + "行电量、占比、单价、扣税率及少回收电能量电费；继续生成时分表和汇总均采用分表自算结果。"
                });
            }
        }

        private static ChongqingStage2Options CreateStagingOptions(
            ChongqingStage2Options source,
            string stagingDirectory)
        {
            var result = new ChongqingStage2Options
            {
                Month = source.Month,
                LedgerPath = source.LedgerPath,
                ProxyTemplateDirectory = source.ProxyTemplateDirectory,
                IntermediaryTemplateDirectory = source.IntermediaryTemplateDirectory,
                RefundTemplateDirectory = source.RefundTemplateDirectory,
                SummaryTemplatePath = source.SummaryTemplatePath,
                OutputDirectory = stagingDirectory,
                OutputSummaryName = source.OutputSummaryName,
                ExpectedPreflightSignature = source.ExpectedPreflightSignature,
                ExpectedInputFingerprint = source.ExpectedInputFingerprint
            };
            result.SummarySubjectDecisions.AddRange(source.SummarySubjectDecisions
                .Where(item => item != null)
                .Select(item => new ChongqingStage2SummarySubjectDecision
                {
                    SettlementKind = item.SettlementKind,
                    Entity = item.Entity,
                    PaymentParty = item.PaymentParty
                }));
            return result;
        }

        private static string CaptureInputFingerprint(
            ChongqingStage2Options options,
            IEnumerable<string> observedOutputPaths)
        {
            return Stage2InputFingerprint.Capture(
                new[]
                {
                    options.LedgerPath,
                    options.SummaryTemplatePath
                },
                new[]
                {
                    options.ProxyTemplateDirectory,
                    options.IntermediaryTemplateDirectory,
                    options.RefundTemplateDirectory
                }.Where(path => !string.IsNullOrWhiteSpace(path)),
                new[]
                {
                    "重庆",
                    options.Month.ToString(CultureInfo.InvariantCulture),
                    Path.GetFullPath(options.OutputDirectory),
                    options.OutputSummaryName ?? string.Empty
                },
                observedOutputPaths);
        }

        private static List<string> BuildObservedOutputPaths(
            ChongqingStage2Options options,
            IEnumerable<ChongqingManagedOutputPlanItem> managedOutputPlan)
        {
            return managedOutputPlan
                .Select(item => item.Path)
                .Concat(new[] { ChongqingStage2SummaryWorkbookWriter.PlanOutputPath(options) })
                .Select(Path.GetFullPath)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static void EnsureGenerationWasAuthorizedByPreflight(
            ChongqingStage2Options options)
        {
            if (string.IsNullOrWhiteSpace(options.ExpectedPreflightSignature)
                || string.IsNullOrWhiteSpace(options.ExpectedInputFingerprint))
            {
                throw new InvalidOperationException(
                    "重庆阶段二正式生成缺少本次预检签名，已拒绝绕过预检；请重新打开预检并确认。" );
            }
        }

        private static void EnsureInputFingerprintUnchanged(
            string expected,
            string actual,
            string phase)
        {
            if (!Stage2InputFingerprint.Matches(expected, actual))
            {
                throw new InvalidOperationException(
                    "重庆阶段二" + phase
                    + "输入文件或模板目录发生变化，本次不会发布；请重新预检并确认。" );
            }
        }

        private static void EnsureGeneratedGroupsMatch(
            ISet<string> expectedKeys,
            IList<GroupSettlementTotal> groups)
        {
            var generatedKeys = groups
                .Select(group => ChongqingStage2Keys.SummaryKey(group.Entity, group.Kind))
                .ToList();
            if (generatedKeys.Count != generatedKeys.Distinct().Count()
                || !expectedKeys.SetEquals(generatedKeys))
            {
                throw new InvalidDataException("重庆阶段二生成的分表主体集合与预检结果不一致。");
            }
        }

        private static void EnsureGeneratedOutputsMatchPlan(
            Stage2BatchWorkspace workspace,
            IEnumerable<GroupSettlementTotal> groups,
            IEnumerable<ChongqingManagedOutputPlanItem> planned)
        {
            var plannedByKey = planned.ToDictionary(
                item => ChongqingStage2Keys.SummaryKey(item.Entity, item.Kind),
                item => Path.GetFullPath(item.Path));
            var generated = groups.ToList();
            if (plannedByKey.Count != generated.Count)
            {
                throw new InvalidDataException("重庆阶段二生成的分表路径集合与预检规划数量不一致。");
            }

            foreach (var group in generated)
            {
                var key = ChongqingStage2Keys.SummaryKey(group.Entity, group.Kind);
                string plannedPath;
                var finalPath = Path.GetFullPath(workspace.GetFinalPath(group.OutputFile));
                if (!plannedByKey.TryGetValue(key, out plannedPath)
                    || !string.Equals(plannedPath, finalPath, StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidDataException(
                        "重庆阶段二分表实际输出路径与预检规划不一致："
                        + group.Kind + " " + group.Entity + "。");
                }
            }
        }

        private static void AddPreflightAuditIssues(
            ChongqingStage2Options options,
            ChongqingStage2PreflightReport preflight,
            IList<ChongqingStage2CheckIssue> auditIssues)
        {
            foreach (var issue in preflight.Issues)
            {
                if (issue.Disposition == Stage2PreflightDisposition.Review
                    || issue.Disposition == Stage2PreflightDisposition.Information)
                {
                    auditIssues.Add(issue);
                    continue;
                }

                if (!issue.RequiresPaymentPartySelection)
                {
                    continue;
                }

                var key = ChongqingStage2Keys.SummaryKey(issue.Entity, issue.SettlementKind);
                var decision = options.SummarySubjectDecisions
                    .Where(item => item != null)
                    .Single(item => ChongqingStage2Keys.SummaryKey(item.Entity, item.SettlementKind) == key);
                auditIssues.Add(new ChongqingStage2CheckIssue
                {
                    Code = Stage2PreflightIssueKinds.PaymentPartyRequired,
                    Disposition = Stage2PreflightDisposition.Information,
                    Severity = "信息",
                    Category = "本次支付方选择",
                    Kind = issue.Kind,
                    SettlementKind = issue.SettlementKind,
                    Owner = issue.Owner,
                    Entity = issue.Entity,
                    TemplateFile = issue.TemplateFile,
                    SheetName = issue.SheetName,
                    PreviousValue = issue.PreviousValue,
                    CurrentValue = "本次选择：" + decision.PaymentParty,
                    Message = "操作员已在本次阶段二预检中明确选择支付方。",
                    Suggestion = "生成后请再核对汇总表支付方和完整收款人字段。"
                });
            }
        }

        private static void RemapOutputReferences(
            Stage2BatchWorkspace workspace,
            IList<GroupSettlementTotal> groups,
            IList<string> warnings,
            IList<ChongqingStage2CheckIssue> auditIssues)
        {
            foreach (var group in groups)
            {
                group.OutputFile = workspace.GetFinalPath(group.OutputFile);
            }

            for (var index = 0; index < warnings.Count; index++)
            {
                warnings[index] = ReplacePathRoot(
                    warnings[index],
                    workspace.StagingDirectory,
                    workspace.OutputDirectory);
            }

            foreach (var issue in auditIssues.Where(item => item != null))
            {
                if (IsStrictlyWithin(workspace.StagingDirectory, issue.TemplateFile))
                {
                    issue.TemplateFile = workspace.GetFinalPath(issue.TemplateFile);
                }
            }
        }

        private static string ReplacePathRoot(string text, string oldRoot, string newRoot)
        {
            if (string.IsNullOrEmpty(text))
            {
                return text;
            }

            var index = text.IndexOf(oldRoot, StringComparison.OrdinalIgnoreCase);
            return index < 0
                ? text
                : text.Substring(0, index) + newRoot + text.Substring(index + oldRoot.Length);
        }

        private static bool IsStrictlyWithin(string root, string candidate)
        {
            if (string.IsNullOrWhiteSpace(candidate))
            {
                return false;
            }

            var prefix = Path.GetFullPath(root).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
                + Path.DirectorySeparatorChar;
            return Path.GetFullPath(candidate).StartsWith(prefix, StringComparison.OrdinalIgnoreCase);
        }

        private static Exception PreserveFailedWorkspace(
            Stage2BatchWorkspace workspace,
            Exception cause,
            string provinceLabel)
        {
            try
            {
                var failedDirectory = workspace.PreserveAsFailed(cause.Message);
                return new InvalidOperationException(
                    provinceLabel + "阶段二本批未完成，没有发布为正式付款结果。未完成内容已保留在："
                    + failedDirectory + "。原因：" + cause.Message,
                    cause);
            }
            catch (Exception preserveException)
            {
                return new InvalidOperationException(
                    provinceLabel + "阶段二本批未完成，且未完成工作区标记失败。原因："
                    + cause.Message + "；标记失败：" + preserveException.Message,
                    new AggregateException(cause, preserveException));
            }
        }

        private static void EnsurePreflightAllowsGeneration(
            ChongqingStage2Options options,
            ChongqingStage2PreflightReport report)
        {
            var evaluation = Stage2PreflightPolicy.Evaluate(report.Issues, options.SummarySubjectDecisions);
            if (evaluation.HasBlockingIssues)
            {
                throw new InvalidOperationException("重庆阶段二预检存在阻断项，请修正台账或模板后重新预检。");
            }

            if (!evaluation.CanContinue)
            {
                throw new InvalidOperationException("重庆阶段二支付方选择尚未完成或存在冲突，请重新预检后再生成。");
            }
        }
    }
}
