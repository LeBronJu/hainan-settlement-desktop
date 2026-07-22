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
    internal sealed class HainanStage2SettlementGenerator
    {
        public HainanStage2Report Generate(HainanStage2Options options)
        {
            EnsureGenerationWasAuthorizedByPreflight(options);
            var initialInputFingerprint = CaptureInputFingerprint(options);
            var snapshot = HainanStage2LedgerReader.ReadSnapshot(options.LedgerPath, options.Month);
            var proxyRows = snapshot.ProxyRows;
            var interRows = snapshot.IntermediaryRows;
            var missingOwners = snapshot.Issues
                .Where(issue => issue.Code == Stage2PreflightIssueKinds.FirstOwnerMissing)
                .Select(issue => "第" + issue.LedgerRow + "行 " + issue.Customer + " 缺少负责人")
                .Distinct()
                .ToList();

            var templateCatalog = HainanStage2TemplateIndex.Build(options.ProxyTemplateDirectory, options.IntermediaryTemplateDirectory);
            var managedOutputPlan = BuildManagedOutputPlan(options, snapshot.SubjectGroups, templateCatalog);
            if (!Stage2InputFingerprint.Matches(
                initialInputFingerprint,
                CaptureInputFingerprint(options)))
            {
                throw InputsChangedException();
            }

            var observedOutputPaths = BuildObservedOutputPaths(options, managedOutputPlan);
            var preflightInputFingerprint = CaptureInputFingerprint(options, observedOutputPaths);
            var preflightIssues = new List<HainanStage2CheckIssue>(snapshot.Issues);
            AddLedgerDifferenceIssues(snapshot, options.Month, preflightIssues);
            preflightIssues.AddRange(BuildPreflightIssues(
                options,
                snapshot.SubjectGroups,
                templateCatalog,
                managedOutputPlan));
            var generationInputFingerprint = CaptureInputFingerprint(options, observedOutputPaths);
            if (!Stage2InputFingerprint.Matches(preflightInputFingerprint, generationInputFingerprint))
            {
                throw InputsChangedException();
            }

            var preflightSignature = Stage2PreflightSignature.Create(
                options.Month,
                snapshot.SubjectGroups.Count,
                preflightIssues);
            if (!Stage2InputFingerprint.Matches(
                    options.ExpectedInputFingerprint,
                    generationInputFingerprint)
                || !Stage2PreflightSignature.Matches(
                    options.ExpectedPreflightSignature,
                    preflightSignature))
            {
                throw InputsChangedException();
            }

            ValidatePreflightPolicy(options, preflightIssues);

            var expectedKeys = new HashSet<string>(snapshot.SubjectGroups.Select(group =>
                HainanStage2ExcelUtil.SummaryKey(group.Entity, group.SettlementKind)));
            var workspace = new Stage2BatchWorkspace(options.OutputDirectory, "海南", options.Month);
            var stagingOptions = CloneOptionsWithOutputDirectory(options, workspace.StagingDirectory);
            try
            {
                var auditIssues = new List<HainanStage2CheckIssue>();
                AddPreflightAuditIssues(options, preflightIssues, auditIssues);
                var totals = HainanStage2SplitWorkbookWriter.BuildSplitFiles(
                    stagingOptions,
                    snapshot.SubjectGroups,
                    templateCatalog);
                EnsureGeneratedGroupsMatch(expectedKeys, totals);

                var warnings = new List<string>();
                var stagedSummaryPath = HainanStage2SummaryWorkbookWriter.BuildSummary(
                    stagingOptions,
                    totals,
                    snapshot.SubjectGroups,
                    warnings);
                HainanStage2SplitWorkbookWriter.VerifyGeneratedSplitWorkbooks(
                    snapshot.SubjectGroups,
                    totals,
                    options.Month);
                HainanStage2SummaryWorkbookWriter.VerifyGeneratedSummary(
                    stagingOptions,
                    totals,
                    snapshot.SubjectGroups,
                    stagedSummaryPath);
                EnsureGeneratedOutputPathsMatchPlan(workspace, managedOutputPlan, totals);

                var stagedSplitPaths = totals.Select(total => total.OutputFile).ToList();
                RemapOutputReferences(workspace, totals, warnings, auditIssues);
                var finalSummaryPath = workspace.GetFinalPath(stagedSummaryPath);
                var finalOptions = CloneOptionsWithOutputDirectory(options, workspace.OutputDirectory);
                var report = HainanStage2ReportWriter.CreateReport(
                    finalOptions,
                    proxyRows,
                    interRows,
                    totals,
                    finalSummaryPath,
                    warnings,
                    missingOwners,
                    auditIssues);
                report.PreflightSignature = preflightSignature;
                report.InputFingerprint = generationInputFingerprint;
                var stagedReportPath = workspace.GetStagingPath(Path.GetFileName(report.ReportPath));
                var stagedWarningsPath = workspace.GetStagingPath(Path.GetFileName(report.GeneratedSummaryReviewPath));
                var stagedValidationPath = workspace.GetStagingPath(Path.GetFileName(report.ValidationReportPath));
                HainanStage2ReportWriter.WriteReport(report, stagedReportPath);
                HainanStage2ReportWriter.WriteWarnings(warnings, stagedWarningsPath);
                HainanStage2ReportWriter.WriteAuditReport(finalOptions, report, stagedValidationPath);

                var reportPaths = new List<string>
                {
                    stagedReportPath,
                    stagedWarningsPath,
                    stagedValidationPath
                };

                Stage2BatchIntegrityVerifier.VerifyFiles(
                    workspace,
                    stagedSplitPaths,
                    stagedSummaryPath,
                    reportPaths);
                EnsureFormalOutputsStillMatchConfirmedState(
                    options,
                    managedOutputPlan,
                    observedOutputPaths,
                    generationInputFingerprint);
                workspace.Publish(() => EnsureFormalOutputsStillMatchConfirmedState(
                    options,
                    managedOutputPlan,
                    observedOutputPaths,
                    generationInputFingerprint));
                return report;
            }
            catch (Exception ex)
            {
                throw PreserveFailedWorkspace(workspace, ex);
            }
        }

        public HainanStage2PreflightReport Analyze(HainanStage2Options options)
        {
            var initialInputFingerprint = CaptureInputFingerprint(options);
            var snapshot = HainanStage2LedgerReader.ReadSnapshot(options.LedgerPath, options.Month);
            var templateCatalog = HainanStage2TemplateIndex.Build(options.ProxyTemplateDirectory, options.IntermediaryTemplateDirectory);
            var managedOutputPlan = BuildManagedOutputPlan(options, snapshot.SubjectGroups, templateCatalog);
            if (!Stage2InputFingerprint.Matches(
                initialInputFingerprint,
                CaptureInputFingerprint(options)))
            {
                throw InputsChangedException();
            }

            var observedOutputPaths = BuildObservedOutputPaths(options, managedOutputPlan);
            var preflightInputFingerprint = CaptureInputFingerprint(options, observedOutputPaths);

            var report = new HainanStage2PreflightReport
            {
                Month = options.Month,
                SubjectCount = snapshot.SubjectGroups.Count
            };
            report.Issues.AddRange(snapshot.Issues);
            AddLedgerDifferenceIssues(snapshot, options.Month, report.Issues);
            report.Issues.AddRange(BuildPreflightIssues(
                options,
                snapshot.SubjectGroups,
                templateCatalog,
                managedOutputPlan));
            var finalInputFingerprint = CaptureInputFingerprint(options, observedOutputPaths);
            if (!Stage2InputFingerprint.Matches(preflightInputFingerprint, finalInputFingerprint))
            {
                throw InputsChangedException();
            }

            report.InputFingerprint = finalInputFingerprint;
            report.PreflightSignature = Stage2PreflightSignature.Create(
                options.Month,
                report.SubjectCount,
                report.Issues);
            return report;
        }

        private static string CaptureInputFingerprint(
            HainanStage2Options options,
            IEnumerable<string> observedOutputPaths = null)
        {
            return Stage2InputFingerprint.Capture(
                new[] { options.LedgerPath, options.SummaryTemplatePath },
                new[] { options.ProxyTemplateDirectory, options.IntermediaryTemplateDirectory },
                new[]
                {
                    "province|海南",
                    "month|" + options.Month.ToString(CultureInfo.InvariantCulture),
                    "output-directory|" + Path.GetFullPath(options.OutputDirectory),
                    "output-summary-name|" + (options.OutputSummaryName ?? string.Empty),
                    "allow-missing-owner|" + options.AllowMissingOwner.ToString()
                },
                observedOutputPaths);
        }

        private static List<string> BuildObservedOutputPaths(
            HainanStage2Options options,
            IEnumerable<ManagedOutputPlanItem> managedOutputPlan)
        {
            return managedOutputPlan
                .Select(item => item.Path)
                .Concat(new[] { HainanStage2SummaryWorkbookWriter.PlanOutputPath(options) })
                .Select(Path.GetFullPath)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static InvalidOperationException InputsChangedException()
        {
            return new InvalidOperationException("海南阶段二输入已变化，请重新预检。");
        }

        private static void EnsureGenerationWasAuthorizedByPreflight(
            HainanStage2Options options)
        {
            if (string.IsNullOrWhiteSpace(options.ExpectedPreflightSignature)
                || string.IsNullOrWhiteSpace(options.ExpectedInputFingerprint))
            {
                throw new InvalidOperationException(
                    "海南阶段二正式生成缺少本次预检签名，已拒绝绕过预检；请重新打开预检并确认。");
            }
        }

        private static void AddLedgerDifferenceIssues(
            HainanStage2LedgerSnapshot snapshot,
            int month,
            IList<HainanStage2CheckIssue> issues)
        {
            foreach (var row in snapshot.ProxyRows)
            {
                var issue = HainanStage2AuditIssueFactory.CreateLedgerDifferenceIssue(
                    row,
                    "代理",
                    null,
                    month + "月");
                if (issue != null)
                {
                    issues.Add(issue);
                }
            }

            foreach (var row in snapshot.IntermediaryRows)
            {
                var issue = HainanStage2AuditIssueFactory.CreateLedgerDifferenceIssue(
                    row,
                    "居间",
                    null,
                    month + "月");
                if (issue != null)
                {
                    issues.Add(issue);
                }
            }
        }

        private static HainanStage2Options CloneOptionsWithOutputDirectory(
            HainanStage2Options source,
            string outputDirectory)
        {
            var result = new HainanStage2Options
            {
                Month = source.Month,
                LedgerPath = source.LedgerPath,
                ProxyTemplateDirectory = source.ProxyTemplateDirectory,
                IntermediaryTemplateDirectory = source.IntermediaryTemplateDirectory,
                SummaryTemplatePath = source.SummaryTemplatePath,
                OutputDirectory = outputDirectory,
                OutputSummaryName = source.OutputSummaryName,
                AllowMissingOwner = source.AllowMissingOwner,
                ExpectedPreflightSignature = source.ExpectedPreflightSignature,
                ExpectedInputFingerprint = source.ExpectedInputFingerprint
            };
            result.SummarySubjectDecisions.AddRange(source.SummarySubjectDecisions
                .Where(item => item != null)
                .Select(item => new HainanStage2SummarySubjectDecision
                {
                    SettlementKind = item.SettlementKind,
                    Entity = item.Entity,
                    PaymentParty = item.PaymentParty
                }));
            return result;
        }

        private static void EnsureGeneratedGroupsMatch(
            ISet<string> expectedKeys,
            IList<GroupSettlementTotal> totals)
        {
            var generatedKeys = totals
                .Select(total => HainanStage2ExcelUtil.SummaryKey(total.Entity, total.Kind))
                .ToList();
            if (generatedKeys.Count != generatedKeys.Distinct().Count()
                || !expectedKeys.SetEquals(generatedKeys))
            {
                throw new InvalidDataException("海南阶段二生成的分表主体集合与预检结果不一致。");
            }
        }

        private static void AddPreflightAuditIssues(
            HainanStage2Options options,
            IEnumerable<HainanStage2CheckIssue> preflightIssues,
            IList<HainanStage2CheckIssue> auditIssues)
        {
            foreach (var issue in preflightIssues)
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

                var key = HainanStage2ExcelUtil.SummaryKey(issue.Entity, issue.SettlementKind);
                var decision = options.SummarySubjectDecisions
                    .Where(item => item != null)
                    .Single(item => HainanStage2ExcelUtil.SummaryKey(item.Entity, item.SettlementKind) == key);
                auditIssues.Add(new HainanStage2CheckIssue
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
            IList<GroupSettlementTotal> totals,
            IList<string> warnings,
            IList<HainanStage2CheckIssue> auditIssues)
        {
            foreach (var total in totals)
            {
                total.OutputFile = workspace.GetFinalPath(total.OutputFile);
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

        private static Exception PreserveFailedWorkspace(Stage2BatchWorkspace workspace, Exception cause)
        {
            try
            {
                var failedDirectory = workspace.PreserveAsFailed(cause.Message);
                return new InvalidOperationException(
                    "海南阶段二本批未完成，没有发布为正式付款结果。未完成内容已保留在："
                    + failedDirectory + "。原因：" + cause.Message,
                    cause);
            }
            catch (Exception preserveException)
            {
                return new InvalidOperationException(
                    "海南阶段二本批未完成，且未完成工作区标记失败。原因："
                    + cause.Message + "；标记失败：" + preserveException.Message,
                    new AggregateException(cause, preserveException));
            }
        }

        private static List<HainanStage2CheckIssue> BuildPreflightIssues(
            HainanStage2Options options,
            IList<HainanStage2SubjectGroup> subjectGroups,
            HainanStage2TemplateCatalog templateCatalog,
            IList<ManagedOutputPlanItem> managedOutputPlan)
        {
            var issues = new List<HainanStage2CheckIssue>(templateCatalog.Issues);
            foreach (var group in subjectGroups
                .OrderBy(item => item.Kind)
                .ThenBy(item => item.FirstLedgerRow))
            {
                var exactCandidates = templateCatalog.ExactCandidates(group.Kind, group.Entity);
                if (exactCandidates.Count > 1)
                {
                    issues.Add(new HainanStage2CheckIssue
                    {
                        Code = Stage2PreflightIssueKinds.DuplicateExactTemplates,
                        Disposition = Stage2PreflightDisposition.Blocker,
                        Severity = "错误",
                        Category = "同一主体匹配到多个上月分表",
                        Kind = group.SettlementKind,
                        SettlementKind = group.SettlementKind,
                        Owner = group.Owner,
                        Entity = group.Entity,
                        LedgerRow = group.FirstLedgerRow,
                        CurrentValue = string.Join("、", exactCandidates.Select(candidate => candidate.Path)),
                        Message = group.SettlementKind + "主体“" + group.Entity + "”匹配到多个同名上月分表，无法确定应继承哪一个完整模板。",
                        Suggestion = "请只保留正确的上月分表，或修正模板 A2 主体名称后重新预检。"
                    });
                    continue;
                }

                if (exactCandidates.Count == 0)
                {
                    var sameKindCandidates = templateCatalog.CandidatesForKind(group.Kind);
                    if (sameKindCandidates.Count != 1)
                    {
                        issues.Add(new HainanStage2CheckIssue
                        {
                            Code = Stage2PreflightIssueKinds.TemplateMissing,
                            Disposition = Stage2PreflightDisposition.Blocker,
                            Severity = "错误",
                            Category = "没有可用的同类型分表模板",
                            Kind = group.SettlementKind,
                            SettlementKind = group.SettlementKind,
                            Owner = group.Owner,
                            Entity = group.Entity,
                            LedgerRow = group.FirstLedgerRow,
                            CurrentValue = string.Join("、", sameKindCandidates.Select(candidate => candidate.Path)),
                            Message = sameKindCandidates.Count == 0
                                ? group.SettlementKind + "主体“" + group.Entity + "”没有同名上月分表，且模板目录中没有可借用的同类型模板。"
                                : group.SettlementKind + "主体“" + group.Entity + "”没有同名上月分表，但找到了多个同类型借用候选，无法可靠确定模板。",
                            Suggestion = sameKindCandidates.Count == 0
                                ? "请补充一个可读取的同类型分表模板后重新预检。"
                                : "请为该主体提供唯一同名模板，或只保留一个明确可借用的同类型模板后重新预检。"
                        });
                    }
                    else
                    {
                        var borrowed = sameKindCandidates[0];
                        issues.Add(new HainanStage2CheckIssue
                        {
                            Code = Stage2PreflightIssueKinds.BorrowedTemplate,
                            Disposition = Stage2PreflightDisposition.Review,
                            Severity = "提示",
                            Category = "新增主体借用同类型分表模板",
                            Kind = group.SettlementKind,
                            SettlementKind = group.SettlementKind,
                            Owner = group.Owner,
                            Entity = group.Entity,
                            LedgerRow = group.FirstLedgerRow,
                            TemplateFile = borrowed.Path,
                            Message = group.SettlementKind + "主体“" + group.Entity + "”未匹配到同名上月分表，将借用同类型模板新建分表。",
                            Suggestion = "请确认这是新增关系；生成后只会保留本月工作表。"
                        });
                    }

                    continue;
                }

                CompareGroupWithTemplate(
                    options.Month,
                    group.Kind,
                    group.Owner,
                    group.Entity,
                    exactCandidates[0].Path,
                    group.Rows,
                    issues);
            }

            AddSummarySubjectPaymentIssues(options, subjectGroups, issues);
            AddManagedOutputIssues(options, managedOutputPlan, issues);
            return issues;
        }

        private static List<ManagedOutputPlanItem> BuildManagedOutputPlan(
            HainanStage2Options options,
            IEnumerable<HainanStage2SubjectGroup> subjectGroups,
            HainanStage2TemplateCatalog templateCatalog)
        {
            return subjectGroups
                .Select(group => new ManagedOutputPlanItem
                {
                    Group = group,
                    Path = HainanStage2SplitWorkbookWriter.PlanOutputPath(options, group, templateCatalog)
                })
                .ToList();
        }

        private static void AddManagedOutputIssues(
            HainanStage2Options options,
            IList<ManagedOutputPlanItem> plans,
            IList<HainanStage2CheckIssue> issues)
        {
            foreach (var conflict in plans
                .GroupBy(plan => plan.Path, StringComparer.OrdinalIgnoreCase)
                .Where(group => group.Count() > 1))
            {
                var first = conflict.First().Group;
                issues.Add(new HainanStage2CheckIssue
                {
                    Code = Stage2PreflightIssueKinds.PlannedOutputPathConflict,
                    Disposition = Stage2PreflightDisposition.Blocker,
                    Severity = "错误",
                    Category = "本批多个主体规划到同一分表路径",
                    Kind = first.SettlementKind,
                    SettlementKind = first.SettlementKind,
                    Owner = first.Owner,
                    Entity = first.Entity,
                    TemplateFile = conflict.Key,
                    CurrentValue = string.Join("、", conflict.Select(plan =>
                        plan.Group.SettlementKind + " " + plan.Group.Entity)),
                    Message = "本批多个汇总身份经文件名安全化或模板文件名继承后，会写入同一正式分表路径：" + conflict.Key,
                    Suggestion = "请调整主体名称或模板文件名，确保每个费用类型+主体有唯一输出文件。"
                });
            }

            var findings = Stage2ManagedOutputInspector.InspectUnexpectedWorkbooks(
                ManagedOutputRoots(options),
                plans.Select(plan => plan.Path),
                options.Month + "月");
            foreach (var finding in findings)
            {
                var isReview = finding.IsPlannedTargetMonth;
                issues.Add(new HainanStage2CheckIssue
                {
                    Code = finding.IsPlannedTargetMonth
                        ? Stage2PreflightIssueKinds.PlannedTargetMonthWorkbook
                        : finding.IsUnreadable
                            ? Stage2PreflightIssueKinds.ManagedOutputUnreadable
                            : Stage2PreflightIssueKinds.UnexpectedTargetMonthWorkbook,
                    Disposition = isReview
                        ? Stage2PreflightDisposition.Review
                        : Stage2PreflightDisposition.Blocker,
                    Severity = isReview ? "复核" : "错误",
                    Category = isReview
                        ? "本批计划分表已含目标月"
                        : finding.IsUnreadable
                            ? "受管输出文件无法安全读取"
                            : "非本批计划分表仍含目标月",
                    TemplateFile = finding.Path,
                    CurrentValue = options.Month + "月",
                    Message = finding.Message,
                    Suggestion = isReview
                        ? "请确认该文件中的人工修改已回填到本次输入模板；继续后将由整包事务覆盖。"
                        : "请先人工核对该文件；程序不会自动删除或移动历史 workbook。"
                });
            }
        }

        private static void EnsureManagedOutputStillSafe(
            HainanStage2Options options,
            IEnumerable<ManagedOutputPlanItem> plans)
        {
            var blockingFindings = Stage2ManagedOutputInspector.InspectUnexpectedWorkbooks(
                    ManagedOutputRoots(options),
                    plans.Select(plan => plan.Path),
                    options.Month + "月")
                .Where(finding => !finding.IsPlannedTargetMonth)
                .ToList();
            if (blockingFindings.Count > 0)
            {
                throw new InvalidDataException(
                    "海南阶段二发布前复检发现正式输出目录已变化，本批不会发布："
                    + string.Join("；", blockingFindings.Select(finding => finding.Message).Take(10)));
            }
        }

        private static void EnsureFormalOutputsStillMatchConfirmedState(
            HainanStage2Options options,
            IEnumerable<ManagedOutputPlanItem> managedOutputPlan,
            IEnumerable<string> observedOutputPaths,
            string confirmedInputFingerprint)
        {
            EnsureManagedOutputStillSafe(options, managedOutputPlan);
            if (!Stage2InputFingerprint.Matches(
                confirmedInputFingerprint,
                CaptureInputFingerprint(options, observedOutputPaths)))
            {
                throw InputsChangedException();
            }
        }

        private static void EnsureGeneratedOutputPathsMatchPlan(
            Stage2BatchWorkspace workspace,
            IEnumerable<ManagedOutputPlanItem> plans,
            IEnumerable<GroupSettlementTotal> totals)
        {
            var plannedByKey = plans.ToDictionary(
                plan => HainanStage2ExcelUtil.SummaryKey(
                    plan.Group.Entity,
                    plan.Group.SettlementKind),
                plan => Path.GetFullPath(plan.Path));
            foreach (var total in totals)
            {
                var key = HainanStage2ExcelUtil.SummaryKey(total.Entity, total.Kind);
                string plannedPath;
                var generatedFinalPath = workspace.GetFinalPath(total.OutputFile);
                if (!plannedByKey.TryGetValue(key, out plannedPath)
                    || !string.Equals(
                        Path.GetFullPath(generatedFinalPath),
                        plannedPath,
                        StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidDataException(
                        "海南阶段二实际生成分表路径与预检计划不一致：" + total.Kind + " " + total.Entity);
                }
            }
        }

        private static IEnumerable<string> ManagedOutputRoots(HainanStage2Options options)
        {
            return new[]
            {
                Path.Combine(options.OutputDirectory, "2026年代理 - 海南"),
                Path.Combine(options.OutputDirectory, "2026年居间 - 海南")
            };
        }

        private static void AddSummarySubjectPaymentIssues(
            HainanStage2Options options,
            IList<HainanStage2SubjectGroup> subjectGroups,
            IList<HainanStage2CheckIssue> issues)
        {
            using (var workbook = new XLWorkbook(options.SummaryTemplatePath))
            {
                var mainCandidates = HainanStage2SummaryWorkbookWriter.FindSummarySheetCandidates(workbook, "main");
                var qingnengCandidates = HainanStage2SummaryWorkbookWriter.FindSummarySheetCandidates(workbook, "qingneng");
                var qinghuiCandidates = HainanStage2SummaryWorkbookWriter.FindSummarySheetCandidates(workbook, "qinghui");
                var roleCandidates = new[]
                {
                    new { Role = "主汇总", Candidates = mainCandidates },
                    new { Role = "清能汇总", Candidates = qingnengCandidates },
                    new { Role = "清辉汇总", Candidates = qinghuiCandidates }
                };
                foreach (var ambiguousRole in roleCandidates.Where(item => item.Candidates.Count > 1))
                {
                    issues.Add(new HainanStage2CheckIssue
                    {
                        Code = Stage2PreflightIssueKinds.SummarySheetAmbiguous,
                        Disposition = Stage2PreflightDisposition.Blocker,
                        Severity = "错误",
                        Category = "汇总模板工作表角色歧义",
                        TemplateFile = options.SummaryTemplatePath,
                        CurrentValue = string.Join("、", ambiguousRole.Candidates),
                        Message = ambiguousRole.Role + "角色同时匹配到多个有效工作表，程序不能静默选择其中一份。",
                        Suggestion = "请只保留一个权威的" + ambiguousRole.Role + "工作表后重新预检。"
                    });
                }

                foreach (var missingPaymentSheet in new[]
                {
                    new { Party = HainanStage2PaymentParties.Qingneng, Candidates = qingnengCandidates },
                    new { Party = HainanStage2PaymentParties.Qinghui, Candidates = qinghuiCandidates }
                }.Where(item => item.Candidates.Count == 0))
                {
                    issues.Add(new HainanStage2CheckIssue
                    {
                        Code = Stage2PreflightIssueKinds.SummaryPaymentSheetMissing,
                        Disposition = Stage2PreflightDisposition.Blocker,
                        Severity = "错误",
                        Category = "汇总模板缺少固定支付方工作表",
                        TemplateFile = options.SummaryTemplatePath,
                        CurrentValue = "缺少" + missingPaymentSheet.Party + "汇总表",
                        Message = "汇总模板缺少" + missingPaymentSheet.Party + "汇总表，无法保证操作员本次支付方选择一定有可靠落点。",
                        Suggestion = "请使用同时包含主汇总、清能汇总和清辉汇总的完整模板后重新预检。"
                    });
                }

                if (roleCandidates.Any(item => item.Candidates.Count > 1))
                {
                    return;
                }

                var mainSheetName = HainanStage2SummaryWorkbookWriter.ResolveSummarySheetName(workbook, "main", true);
                var qingnengSheetName = qingnengCandidates.SingleOrDefault();
                var qinghuiSheetName = qinghuiCandidates.SingleOrDefault();
                var selectedSheetNames = new[] { mainSheetName, qingnengSheetName, qinghuiSheetName }
                    .Where(name => !string.IsNullOrWhiteSpace(name))
                    .Distinct(StringComparer.Ordinal)
                    .ToList();
                foreach (var sheetName in selectedSheetNames)
                {
                    foreach (var invalid in HainanStage2SummaryWorkbookWriter.FindInvalidSummaryKindRows(workbook.Worksheet(sheetName)))
                    {
                        issues.Add(new HainanStage2CheckIssue
                        {
                            Code = Stage2PreflightIssueKinds.SummarySubjectKindInvalid,
                            Disposition = Stage2PreflightDisposition.Blocker,
                            Severity = "错误",
                            Category = "汇总表费用类型无效",
                            Kind = invalid.Kind,
                            SettlementKind = invalid.Kind,
                            Entity = invalid.Entity,
                            TemplateFile = options.SummaryTemplatePath,
                            SheetName = sheetName,
                            CurrentValue = string.IsNullOrWhiteSpace(invalid.Kind) ? "空白" : invalid.Kind,
                            Message = sheetName + "第" + invalid.Row + "行主体“" + invalid.Entity + "”的费用类型不是代理费或居间费，程序不会重写该行。",
                            Suggestion = "请修正 C 列费用类型，或将非结算备注移出汇总数据区后重新预检。"
                        });
                    }
                }

                if (issues.Any(issue => issue.Code == Stage2PreflightIssueKinds.SummarySubjectKindInvalid))
                {
                    return;
                }

                var mainRows = HainanStage2SummaryWorkbookWriter.ReadSummaryMeta(workbook.Worksheet(mainSheetName)).ToList();
                var sourceRows = HainanStage2SummaryWorkbookWriter.ReadSummarySources(workbook).ToList();
                var mainKeys = new HashSet<string>(mainRows.Select(SummaryKey));
                var qingnengRows = string.IsNullOrWhiteSpace(qingnengSheetName)
                    ? new List<HainanStage2SummaryMetaRow>()
                    : HainanStage2SummaryWorkbookWriter.ReadSummaryMeta(workbook.Worksheet(qingnengSheetName)).ToList();
                var qinghuiRows = string.IsNullOrWhiteSpace(qinghuiSheetName)
                    ? new List<HainanStage2SummaryMetaRow>()
                    : HainanStage2SummaryWorkbookWriter.ReadSummaryMeta(workbook.Worksheet(qinghuiSheetName)).ToList();
                var qingnengByKey = qingnengRows.GroupBy(SummaryKey).ToDictionary(group => group.Key, group => group.ToList());
                var qinghuiByKey = qinghuiRows.GroupBy(SummaryKey).ToDictionary(group => group.Key, group => group.ToList());
                var dualMembershipKeys = new HashSet<string>(qingnengByKey.Keys.Intersect(qinghuiByKey.Keys));
                var membershipPartyByKey = new Dictionary<string, string>();
                foreach (var key in qingnengByKey.Keys.Where(key => !dualMembershipKeys.Contains(key)))
                {
                    membershipPartyByKey[key] = HainanStage2PaymentParties.Qingneng;
                }
                foreach (var key in qinghuiByKey.Keys.Where(key => !dualMembershipKeys.Contains(key)))
                {
                    membershipPartyByKey[key] = HainanStage2PaymentParties.Qinghui;
                }

                foreach (var conflictingKey in dualMembershipKeys)
                {
                    var first = qingnengByKey[conflictingKey][0];
                    string forcedParty;
                    if (HainanStage2ExcelUtil.TryGetPaymentPartyOverride(
                        first.Entity,
                        first.Kind,
                        options.Month,
                        out forcedParty))
                    {
                        membershipPartyByKey[conflictingKey] = forcedParty;
                        continue;
                    }

                    issues.Add(new HainanStage2CheckIssue
                    {
                        Code = Stage2PreflightIssueKinds.ConflictingPaymentParties,
                        Disposition = Stage2PreflightDisposition.Blocker,
                        Severity = "错误",
                        Category = "同一主体同时出现在两个支付方工作表",
                        Kind = first.Kind,
                        SettlementKind = first.Kind,
                        Entity = first.Entity,
                        TemplateFile = options.SummaryTemplatePath,
                        CurrentValue = qingnengSheetName + "第"
                            + string.Join("、", qingnengByKey[conflictingKey].Select(row => row.Row)) + "行；"
                            + qinghuiSheetName + "第"
                            + string.Join("、", qinghuiByKey[conflictingKey].Select(row => row.Row)) + "行",
                        Message = first.Kind + "主体“" + first.Entity + "”同时出现在清能和清辉汇总表，即使单元格支付方文本相同或为空，也无法证明最终归属。",
                        Suggestion = "请只在正确的一个支付方工作表保留该主体后重新预检。"
                    });
                }

                foreach (var sheetName in new[] { mainSheetName, qingnengSheetName, qinghuiSheetName }
                    .Where(name => !string.IsNullOrWhiteSpace(name))
                    .Distinct(StringComparer.Ordinal))
                {
                    var monthBlocks = HainanStage2SummaryWorkbookWriter.FindSummaryMonthBlocks(
                        workbook.Worksheet(sheetName),
                        options.Month);
                    var monthHeaders = HainanStage2SummaryWorkbookWriter.FindSummaryMonthHeaderColumns(
                        workbook.Worksheet(sheetName),
                        options.Month);
                    if (monthHeaders.Count == 0)
                    {
                        continue;
                    }

                    var isAmbiguous = monthHeaders.Count != 1 || monthBlocks.Count != 1;
                    issues.Add(new HainanStage2CheckIssue
                    {
                        Code = Stage2PreflightIssueKinds.SummaryTargetMonthAlreadyExists,
                        Disposition = isAmbiguous
                            ? Stage2PreflightDisposition.Blocker
                            : Stage2PreflightDisposition.Review,
                        Severity = isAmbiguous ? "错误" : "复核",
                        Category = isAmbiguous ? "汇总表目标月块重复" : "汇总表已有目标月块",
                        TemplateFile = options.SummaryTemplatePath,
                        SheetName = sheetName,
                        CurrentValue = string.Join("、", monthHeaders.Select(column => "第" + column + "列")),
                        Message = isAmbiguous
                            ? sheetName + "中找到多个" + options.Month + "月区块，无法确定应重写哪一个。"
                            : sheetName + "中已有可唯一识别的" + options.Month + "月区块，本次将在保留结构的前提下安全重写。",
                        Suggestion = isAmbiguous
                            ? "请只保留一个正确的目标月区块后重新预检。"
                            : "请确认这是本月重新生成，生成后复核该月金额。"
                    });
                }

                foreach (var orphanGroup in sourceRows
                    .Where(row => row.SheetName != mainSheetName && !mainKeys.Contains(SummaryKey(row)))
                    .GroupBy(row => row.SheetName + "|" + SummaryKey(row)))
                {
                    var orphan = orphanGroup.First();
                    issues.Add(new HainanStage2CheckIssue
                    {
                        Code = Stage2PreflightIssueKinds.SummaryOrphanSubject,
                        Disposition = Stage2PreflightDisposition.Blocker,
                        Severity = "错误",
                        Category = "支付方汇总表存在主表外孤立主体",
                        Kind = orphan.Kind,
                        SettlementKind = orphan.Kind,
                        Entity = orphan.Entity,
                        TemplateFile = options.SummaryTemplatePath,
                        SheetName = orphan.SheetName,
                        CurrentValue = string.Join("、", orphanGroup.Select(row => "第" + row.Row + "行")),
                        Message = orphan.SheetName + "中的" + orphan.Kind + "主体“" + orphan.Entity + "”在主汇总表中不存在，无法证明长期字段的权威来源。",
                        Suggestion = "请在主汇总表补全该主体，或从支付方汇总表删除错误孤立行后重新预检。"
                    });
                }

                foreach (var duplicate in sourceRows
                    .GroupBy(row => row.SheetName + "|" + SummaryKey(row))
                    .Where(group => group.Count() > 1))
                {
                    var first = duplicate.First();
                    issues.Add(new HainanStage2CheckIssue
                    {
                        Code = Stage2PreflightIssueKinds.DuplicateSummarySubject,
                        Disposition = Stage2PreflightDisposition.Blocker,
                        Severity = "错误",
                        Category = "汇总表存在重复主体",
                        Kind = first.Kind,
                        SettlementKind = first.Kind,
                        Entity = first.Entity,
                        TemplateFile = options.SummaryTemplatePath,
                        SheetName = first.SheetName,
                        CurrentValue = string.Join("、", duplicate.Select(row => "第" + row.Row + "行")),
                        Message = first.SheetName + "中" + first.Kind + "主体“" + first.Entity + "”出现多次，无法可靠继承长期字段。",
                        Suggestion = "请合并或删除重复汇总行后重新预检。"
                    });
                }

                foreach (var sourceGroup in sourceRows.GroupBy(SummaryKey))
                {
                    var first = sourceGroup.First();
                    var payees = sourceGroup
                        .Select(row => Stage2OpaqueText.NormalizeForComparison(row.Payee))
                        .Where(value => !string.IsNullOrWhiteSpace(value))
                        .Distinct(StringComparer.Ordinal)
                        .ToList();
                    if (payees.Count > 1)
                    {
                        issues.Add(new HainanStage2CheckIssue
                        {
                            Code = Stage2PreflightIssueKinds.ConflictingPayees,
                            Disposition = Stage2PreflightDisposition.Blocker,
                            Severity = "错误",
                            Category = "完整收款人字段冲突",
                            Kind = first.Kind,
                            SettlementKind = first.Kind,
                            Entity = first.Entity,
                            TemplateFile = options.SummaryTemplatePath,
                            CurrentValue = string.Join("；", sourceGroup
                                .Where(row => !string.IsNullOrWhiteSpace(Stage2OpaqueText.NormalizeForComparison(row.Payee)))
                                .Select(row => row.SheetName + "=" + Stage2OpaqueText.NormalizeForComparison(row.Payee))),
                            Message = first.Kind + "主体“" + first.Entity + "”从汇总表可靠来源读取到多个不同的完整收款人文本，无法确定最终应继承哪个单元格值。",
                            Suggestion = "请统一各汇总 sheet 的完整收款人单元格；程序不会解析或拆分其中的人名。"
                        });
                    }
                    else if (payees.Count == 1
                        && mainKeys.Contains(sourceGroup.Key)
                        && sourceGroup.Any(row =>
                            string.IsNullOrWhiteSpace(Stage2OpaqueText.NormalizeForComparison(row.Payee))))
                    {
                        var canonicalSource = sourceGroup.First(row =>
                            Stage2OpaqueText.NormalizeForComparison(row.Payee) == payees[0]);
                        var blankSources = sourceGroup
                            .Where(row => string.IsNullOrWhiteSpace(Stage2OpaqueText.NormalizeForComparison(row.Payee)))
                            .Select(row => row.SheetName)
                            .Distinct(StringComparer.Ordinal)
                            .ToList();
                        issues.Add(new HainanStage2CheckIssue
                        {
                            Code = Stage2PreflightIssueKinds.PayeeSourceMissing,
                            Disposition = Stage2PreflightDisposition.Review,
                            Severity = "提示",
                            Category = "汇总表可靠来源收款人缺失",
                            Kind = first.Kind,
                            SettlementKind = first.Kind,
                            Entity = first.Entity,
                            TemplateFile = options.SummaryTemplatePath,
                            SheetName = canonicalSource.SheetName,
                            PreviousValue = "空白来源：" + string.Join("、", blankSources),
                            CurrentValue = canonicalSource.Payee,
                            Message = first.Kind + "主体“" + first.Entity + "”在“" + string.Join("、", blankSources)
                                + "”的完整收款人为空，将采用“" + canonicalSource.SheetName + "”中的唯一非空完整文本。",
                            Suggestion = "程序会原样回填完整单元格文本，不解析其中姓名；请生成后复核。"
                        });
                    }
                    else if (payees.Count == 0 && mainKeys.Contains(sourceGroup.Key))
                    {
                        issues.Add(new HainanStage2CheckIssue
                        {
                            Code = Stage2PreflightIssueKinds.PayeeSourceMissing,
                            Disposition = Stage2PreflightDisposition.Review,
                            Severity = "复核",
                            Category = "存量主体收款人仍为空白",
                            Kind = first.Kind,
                            SettlementKind = first.Kind,
                            Entity = first.Entity,
                            TemplateFile = options.SummaryTemplatePath,
                            SheetName = mainSheetName,
                            CurrentValue = "保持空白",
                            Message = first.Kind + "主体“" + first.Entity + "”在所有可靠汇总来源中的完整收款人都为空，本次不会臆造内容。",
                            Suggestion = "允许继续并保持空白；请操作员在付款前人工补全并复核。"
                        });
                    }

                    string overrideParty;
                    var hasOverride = HainanStage2ExcelUtil.TryGetPaymentPartyOverride(first.Entity, first.Kind, options.Month, out overrideParty);
                    var parties = sourceGroup
                        .Select(row => Stage2OpaqueText.NormalizeForComparison(row.PaymentParty))
                        .Where(value => !string.IsNullOrWhiteSpace(value))
                        .Distinct(StringComparer.Ordinal)
                        .ToList();
                    if (hasOverride)
                    {
                        continue;
                    }

                    if (dualMembershipKeys.Contains(sourceGroup.Key))
                    {
                        continue;
                    }

                    string membershipParty;
                    if (membershipPartyByKey.TryGetValue(sourceGroup.Key, out membershipParty))
                    {
                        var conflictingFields = sourceGroup
                            .Select(row => new
                            {
                                row.SheetName,
                                Value = Stage2OpaqueText.NormalizeForComparison(row.PaymentParty)
                            })
                            .Where(item => !string.IsNullOrWhiteSpace(item.Value)
                                && item.Value != membershipParty)
                            .ToList();
                        if (conflictingFields.Count > 0)
                        {
                            issues.Add(new HainanStage2CheckIssue
                            {
                                Code = Stage2PreflightIssueKinds.ConflictingPaymentParties,
                                Disposition = Stage2PreflightDisposition.Blocker,
                                Severity = "错误",
                                Category = "支付方工作表所属与字段冲突",
                                Kind = first.Kind,
                                SettlementKind = first.Kind,
                                Entity = first.Entity,
                                TemplateFile = options.SummaryTemplatePath,
                                CurrentValue = "工作表所属=" + membershipParty + "；"
                                    + string.Join("；", conflictingFields.Select(item => item.SheetName + "=" + item.Value)),
                                Message = first.Kind + "主体“" + first.Entity + "”只出现在" + membershipParty
                                    + "汇总表，但主表或支付方表的非空支付方字段与该所属冲突。",
                                Suggestion = "请统一工作表所属和支付方字段后重新预检。"
                            });
                        }

                        continue;
                    }

                    if (parties.Count > 1)
                    {
                        issues.Add(new HainanStage2CheckIssue
                        {
                            Code = Stage2PreflightIssueKinds.ConflictingPaymentParties,
                            Disposition = Stage2PreflightDisposition.Blocker,
                            Severity = "错误",
                            Category = "支付方字段冲突",
                            Kind = first.Kind,
                            SettlementKind = first.Kind,
                            Entity = first.Entity,
                            TemplateFile = options.SummaryTemplatePath,
                            CurrentValue = string.Join("；", sourceGroup
                                .Where(row => !string.IsNullOrWhiteSpace(Stage2OpaqueText.NormalizeForComparison(row.PaymentParty)))
                                .Select(row => row.SheetName + "=" + Stage2OpaqueText.NormalizeForComparison(row.PaymentParty))),
                            Message = first.Kind + "主体“" + first.Entity + "”从汇总表可靠来源读取到多个不同支付方，无法确定应归清能还是清辉。",
                            Suggestion = "请先统一汇总表中的支付方字段后重新预检。"
                        });
                    }
                    else if (mainKeys.Contains(sourceGroup.Key)
                        && (parties.Count != 1 || !HainanStage2PaymentParties.Supported.Contains(parties[0])))
                    {
                        AddPaymentPartyRequirement(
                            issues,
                            options,
                            mainSheetName,
                            first.Kind,
                            first.Entity,
                            null,
                            "存量汇总主体支付方选择",
                            "存量汇总主体支付方为空或不是受支持的单一值，不能再默认归清辉。");
                    }
                }

                foreach (var subject in subjectGroups)
                {
                    var key = HainanStage2ExcelUtil.SummaryKey(subject.Entity, subject.SettlementKind);
                    if (!mainKeys.Contains(key))
                    {
                        issues.Add(new HainanStage2CheckIssue
                        {
                            Code = Stage2PreflightIssueKinds.NewSummarySubject,
                            Disposition = Stage2PreflightDisposition.Review,
                            Severity = "提示",
                            Category = "新增汇总主体默认资料",
                            Kind = subject.SettlementKind,
                            SettlementKind = subject.SettlementKind,
                            Owner = subject.Owner,
                            Entity = subject.Entity,
                            LedgerRow = subject.FirstLedgerRow,
                            TemplateFile = options.SummaryTemplatePath,
                            SheetName = mainSheetName,
                            CurrentValue = "不委托；收款人=" + subject.Entity + "；发票票种=平台；扣税率=" + subject.TaxRate.ToString("0.####%") + "；合计税率=13%",
                            Message = "汇总表将新增" + subject.SettlementKind + "主体“" + subject.Entity + "”，并写入已确认的默认资料和税率公式。",
                            Suggestion = "生成后请检查完整收款人、是否委托、发票票种、税率、负责人及其它长期字段。"
                        });

                        string paymentParty;
                        if (!HainanStage2ExcelUtil.TryGetPaymentPartyOverride(subject.Entity, subject.SettlementKind, options.Month, out paymentParty))
                        {
                            AddPaymentPartyRequirement(
                                issues,
                                options,
                                mainSheetName,
                                subject.SettlementKind,
                                subject.Entity,
                                subject.Owner,
                                "新增汇总主体支付方选择",
                                "新增汇总主体没有可继承的支付方。");
                        }
                    }

                    var matchingSources = sourceRows.Where(row => SummaryKey(row) == key).ToList();
                    var inheritedPayee = matchingSources
                        .Select(row => row.Payee)
                        .FirstOrDefault(value => !string.IsNullOrWhiteSpace(Stage2OpaqueText.NormalizeForComparison(value)));
                    string inheritedParty;
                    if (!HainanStage2ExcelUtil.TryGetPaymentPartyOverride(subject.Entity, subject.SettlementKind, options.Month, out inheritedParty))
                    {
                        if (!membershipPartyByKey.TryGetValue(key, out inheritedParty))
                        {
                            var supportedParties = matchingSources
                                .Select(row => Stage2OpaqueText.NormalizeForComparison(row.PaymentParty))
                                .Where(value => HainanStage2PaymentParties.Supported.Contains(value))
                                .Distinct(StringComparer.Ordinal)
                                .ToList();
                            inheritedParty = supportedParties.Count == 1 ? supportedParties[0] : null;
                        }
                    }

                    foreach (var issue in issues.Where(issue =>
                        issue.Code == Stage2PreflightIssueKinds.NewCustomer
                        && HainanStage2ExcelUtil.SummaryKey(issue.Entity, issue.SettlementKind) == key))
                    {
                        issue.CurrentValue = "继承收款人：" + (string.IsNullOrWhiteSpace(inheritedPayee) ? "空白" : inheritedPayee)
                            + "；支付方：" + (string.IsNullOrWhiteSpace(inheritedParty) ? "待选择" : inheritedParty);
                        issue.Suggestion = "这是存量主体下的新增客户；将继承该主体完整收款人和支付方。请同时检查客户关系参数。";
                    }
                }

            }
        }

        private static string SummaryKey(HainanStage2SummaryMetaRow row)
        {
            return HainanStage2ExcelUtil.SummaryKey(row.Entity, row.Kind);
        }

        private static void AddPaymentPartyRequirement(
            IList<HainanStage2CheckIssue> issues,
            HainanStage2Options options,
            string sheetName,
            string kind,
            string entity,
            string owner,
            string category,
            string reason)
        {
            var key = HainanStage2ExcelUtil.SummaryKey(entity, kind);
            if (issues.Any(existingIssue =>
                existingIssue.Code == Stage2PreflightIssueKinds.PaymentPartyRequired
                && HainanStage2ExcelUtil.SummaryKey(existingIssue.Entity, existingIssue.SettlementKind) == key))
            {
                return;
            }

            var issue = new HainanStage2CheckIssue
            {
                Code = Stage2PreflightIssueKinds.PaymentPartyRequired,
                Disposition = Stage2PreflightDisposition.RequiredDecision,
                Severity = "确认",
                Category = category,
                Kind = kind,
                SettlementKind = kind,
                Owner = owner,
                Entity = entity,
                TemplateFile = options.SummaryTemplatePath,
                SheetName = sheetName,
                Message = reason,
                Suggestion = "请选择本月输出使用的支付方清能或清辉；本次选择只写入输出副本。",
                RequiresPaymentPartySelection = true
            };
            issue.AvailablePaymentParties.AddRange(HainanStage2PaymentParties.Supported);
            issues.Add(issue);
        }

        private static void ValidatePreflightPolicy(
            HainanStage2Options options,
            IList<HainanStage2CheckIssue> issues)
        {
            var evaluation = Stage2PreflightPolicy.Evaluate(issues, options.SummarySubjectDecisions);
            if (!evaluation.CanContinue)
            {
                var messages = issues
                    .Where(issue => issue.BlocksGeneration)
                    .Select(issue => issue.Message)
                    .Concat(evaluation.InvalidDefinitions)
                    .Concat(evaluation.DecisionResolutions
                        .Where(item => item.Status != Stage2PaymentPartyDecisionStatus.Resolved)
                        .Select(item => item.SettlementKind + " " + item.Entity + "：" + item.Message))
                    .Where(message => !string.IsNullOrWhiteSpace(message))
                    .Distinct()
                    .Take(10)
                    .ToList();
                var prefix = evaluation.HasOutstandingRequiredDecisions || evaluation.HasInvalidDecisions
                    ? "海南阶段二新增汇总主体支付方未选择或存量支付方待处理，未生成任何正式文件："
                    : "海南阶段二预检未通过，未生成任何正式文件：";
                throw new InvalidOperationException(prefix + string.Join("；", messages));
            }
        }

        private static void CompareGroupWithTemplate(
            int month,
            string kind,
            string owner,
            string entity,
            string templatePath,
            IList<HainanStage2DetailSettlementRow> currentRows,
            IList<HainanStage2CheckIssue> issues)
        {
            try
            {
                using (var workbook = new XLWorkbook(templatePath))
                {
                    var previousSheet = HainanStage2ExcelUtil.PreviousMonthSheet(workbook, month, month + "月") ?? HainanStage2ExcelUtil.LastMonthSheet(workbook);
                    var previousRows = ReadPreviousDetails(previousSheet);
                    var currentCustomers = new HashSet<string>(currentRows.Select(row => HainanStage2ExcelUtil.NormalizeName(row.Customer)));
                    foreach (var row in currentRows)
                    {
                        PreviousDetailRow previous;
                        if (!previousRows.TryGetValue(HainanStage2ExcelUtil.NormalizeName(row.Customer), out previous))
                        {
                            issues.Add(new HainanStage2CheckIssue
                            {
                                Code = Stage2PreflightIssueKinds.NewCustomer,
                                Disposition = Stage2PreflightDisposition.Review,
                                Severity = "提示",
                                Category = "客户本月新增到分表",
                                Kind = kind + "费",
                                SettlementKind = kind + "费",
                                Customer = row.Customer,
                                Owner = owner,
                                Entity = entity,
                                LedgerRow = row.LedgerRow,
                                TemplateFile = templatePath,
                                SheetName = previousSheet.Name,
                                Message = kind + "费主体“" + entity + "”下的客户“" + row.Customer + "”在上月分表未找到。",
                                Suggestion = "如果这是本月新增客户，可以继续；否则请检查台账客户名称和上月分表客户名称是否一致。"
                            });
                            continue;
                        }

                        AddValueChangeIssue(issues, kind, owner, entity, row, previous, templatePath, "电量比例", previous.Ratio, row.Ratio);
                        AddValueChangeIssue(issues, kind, owner, entity, row, previous, templatePath, "利润单价", previous.UnitPrice, row.UnitPrice);
                        AddValueChangeIssue(issues, kind, owner, entity, row, previous, templatePath, "税率", previous.TaxRate, row.TaxRate);
                    }

                    foreach (var previous in previousRows.Values.Where(row => !currentCustomers.Contains(HainanStage2ExcelUtil.NormalizeName(row.Customer))))
                    {
                        issues.Add(new HainanStage2CheckIssue
                        {
                            Code = Stage2PreflightIssueKinds.PreviousTemplateCustomerMissing,
                            Disposition = Stage2PreflightDisposition.Review,
                            Severity = "提示",
                            Category = "上月分表存在本月台账外明细行",
                            Kind = kind + "费",
                            SettlementKind = kind + "费",
                            Customer = previous.Customer,
                            Owner = owner,
                            Entity = entity,
                            TemplateFile = templatePath,
                            SheetName = previous.SheetName,
                            PreviousValue = "上月分表第" + previous.Row + "行",
                            CurrentValue = "本月台账无匹配客户",
                            Message = kind + "费主体“" + entity + "”的上月分表第" + previous.Row + "行“" + previous.Customer + "”未在本月台账中匹配到。",
                            Suggestion = "程序生成本月分表时不会继承该行；如果本月仍需退补或补扣，请生成后手动调整分表和汇总表。"
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                issues.Add(new HainanStage2CheckIssue
                {
                    Code = Stage2PreflightIssueKinds.TemplateUnreadable,
                    Disposition = Stage2PreflightDisposition.Blocker,
                    Severity = "错误",
                    Category = "上月分表预检失败",
                    Kind = kind + "费",
                    SettlementKind = kind + "费",
                    Owner = owner,
                    Entity = entity,
                    TemplateFile = templatePath,
                    Message = "读取上月分表模板失败：" + ex.Message,
                    Suggestion = "请确认该分表文件没有损坏，且包含上月明细和合计行；修复后重新预检。"
                });
            }
        }

        private static Dictionary<string, PreviousDetailRow> ReadPreviousDetails(IXLWorksheet worksheet)
        {
            var result = new Dictionary<string, PreviousDetailRow>();
            var totalRow = HainanStage2ExcelUtil.FindTotalRow(worksheet, HainanStage2ExcelUtil.DataStartRow);
            for (var row = HainanStage2ExcelUtil.DataStartRow; row < totalRow; row++)
            {
                var customer = TextUtil.S(worksheet.Cell(row, 2).GetFormattedString());
                if (string.IsNullOrWhiteSpace(customer))
                {
                    continue;
                }

                var key = HainanStage2ExcelUtil.NormalizeName(customer);
                if (result.ContainsKey(key))
                {
                    continue;
                }

                result[key] = new PreviousDetailRow
                {
                    Customer = customer,
                    Row = row,
                    SheetName = worksheet.Name,
                    Ratio = HainanStage2ExcelUtil.GetNumeric(worksheet, row, 10),
                    UnitPrice = HainanStage2ExcelUtil.GetNumeric(worksheet, row, 11),
                    TaxRate = HainanStage2ExcelUtil.GetNumeric(worksheet, row, 17)
                };
            }

            return result;
        }

        private static void AddValueChangeIssue(
            IList<HainanStage2CheckIssue> issues,
            string kind,
            string owner,
            string entity,
            HainanStage2DetailSettlementRow current,
            PreviousDetailRow previous,
            string templatePath,
            string fieldName,
            double previousValue,
            double currentValue)
        {
            if (Math.Abs(previousValue - currentValue) <= Stage2SettlementCalculator.AmountTolerance)
            {
                return;
            }

            issues.Add(new HainanStage2CheckIssue
            {
                Code = Stage2PreflightIssueKinds.RelationshipValueChanged,
                Disposition = Stage2PreflightDisposition.Review,
                Severity = "确认",
                Category = "关键字段较上月变化",
                Kind = kind + "费",
                SettlementKind = kind + "费",
                Customer = current.Customer,
                Owner = owner,
                Entity = entity,
                LedgerRow = current.LedgerRow,
                TemplateFile = templatePath,
                SheetName = previous.SheetName,
                PreviousValue = fieldName + " 上月：" + Stage2SettlementCalculator.FormatAmount(previousValue),
                CurrentValue = fieldName + " 本月：" + Stage2SettlementCalculator.FormatAmount(currentValue),
                Message = kind + "费主体“" + entity + "”下的客户“" + current.Customer + "”" + fieldName + "发生变化（上月分表第" + previous.Row + "行，本月台账第" + current.LedgerRow + "行）。",
                Suggestion = "如果这是本月台账更新后的正常变化，请继续；否则请回到台账检查该客户的" + fieldName + "。"
            });
        }

        private sealed class ManagedOutputPlanItem
        {
            public HainanStage2SubjectGroup Group { get; set; }
            public string Path { get; set; }
        }

        private sealed class PreviousDetailRow
        {
            public int Row { get; set; }
            public string Customer { get; set; }
            public string SheetName { get; set; }
            public double Ratio { get; set; }
            public double UnitPrice { get; set; }
            public double TaxRate { get; set; }
        }
    }
}
