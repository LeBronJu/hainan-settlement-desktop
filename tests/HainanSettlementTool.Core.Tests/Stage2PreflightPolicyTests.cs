using System.Collections.Generic;
using HainanSettlementTool.Core.Models;
using HainanSettlementTool.Core.Services;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace HainanSettlementTool.Core.Tests
{
    [TestClass]
    public sealed class Stage2PreflightPolicyTests
    {
        [TestMethod]
        public void ReportBlocksGenerationWhenAnyIssueIsBlocker()
        {
            var report = new HainanStage2PreflightReport();
            report.Issues.Add(new HainanStage2CheckIssue
            {
                Disposition = Stage2PreflightDisposition.Review
            });
            report.Issues.Add(new HainanStage2CheckIssue
            {
                Disposition = Stage2PreflightDisposition.Blocker
            });

            Assert.IsTrue(report.HasBlockingIssues);
            Assert.IsFalse(report.CanGenerate);
            Assert.AreEqual(1, report.BlockerCount);
            Assert.AreEqual(1, report.ReviewCount);
        }

        [TestMethod]
        public void BlockingIssueCannotGenerateEvenWhenReviewIsConfirmed()
        {
            var issue = CreateIssue(Stage2PreflightDisposition.Blocker);

            var evaluation = Stage2PreflightPolicy.Evaluate(
                new[] { issue },
                new IStage2PaymentPartyDecision[0]);

            Assert.IsFalse(evaluation.CanGenerate(true));
        }

        [TestMethod]
        public void InformationIssueAllowsGenerationWithoutConfirmation()
        {
            var issue = CreateIssue(Stage2PreflightDisposition.Information);

            var evaluation = Stage2PreflightPolicy.Evaluate(
                new[] { issue },
                new IStage2PaymentPartyDecision[0]);

            Assert.IsTrue(evaluation.CanGenerate(false));
            Assert.IsFalse(evaluation.RequiresConfirmation);
        }

        [TestMethod]
        public void ReviewIssueRequiresExplicitConfirmation()
        {
            var issue = CreateIssue(Stage2PreflightDisposition.Review);

            var evaluation = Stage2PreflightPolicy.Evaluate(
                new[] { issue },
                new IStage2PaymentPartyDecision[0]);

            Assert.IsFalse(evaluation.CanGenerate(false));
            Assert.IsTrue(evaluation.CanGenerate(true));
        }

        [TestMethod]
        public void RequiredDecisionIsNotBlockerButCannotGenerateWithoutResolution()
        {
            var report = new ChongqingStage2PreflightReport();
            report.Issues.Add(new ChongqingStage2CheckIssue
            {
                Disposition = Stage2PreflightDisposition.RequiredDecision,
                RequiresPaymentPartySelection = true
            });

            Assert.IsFalse(report.HasBlockingIssues);
            Assert.IsTrue(report.HasRequiredDecisions);
            Assert.IsFalse(report.CanGenerate);
            Assert.AreEqual(1, report.RequiredDecisionCount);
        }

        [TestMethod]
        public void LegacyPaymentPartyFlagStillCountsAsRequiredDecision()
        {
            var report = new HainanStage2PreflightReport();
            report.Issues.Add(new HainanStage2CheckIssue
            {
                Disposition = Stage2PreflightDisposition.Review,
                RequiresPaymentPartySelection = true
            });

            Assert.IsTrue(report.HasRequiredDecisions);
        }

        [TestMethod]
        public void PaymentPartyRequirementIsOutstandingUntilMatchingDecisionExists()
        {
            var requirement = CreatePaymentPartyRequirement("代理费", "同名主体");

            var outstanding = Stage2PreflightPolicy.Evaluate(
                new[] { requirement },
                new IStage2PaymentPartyDecision[0]);
            var resolved = Stage2PreflightPolicy.Evaluate(
                new[] { requirement },
                new IStage2PaymentPartyDecision[]
                {
                    new HainanStage2SummarySubjectDecision
                    {
                        SettlementKind = "代理费",
                        Entity = "同名主体",
                        PaymentParty = HainanStage2PaymentParties.Qingneng
                    }
                });

            Assert.IsTrue(outstanding.HasOutstandingRequiredDecisions);
            Assert.IsFalse(outstanding.CanContinue);
            Assert.IsFalse(resolved.HasOutstandingRequiredDecisions);
            Assert.IsTrue(resolved.CanContinue);
        }

        [TestMethod]
        public void ProxyDecisionDoesNotResolveIntermediaryRequirementForSameEntity()
        {
            var requirement = CreatePaymentPartyRequirement("居间费", "同名主体");
            var decision = new HainanStage2SummarySubjectDecision
            {
                SettlementKind = "代理费",
                Entity = "同名主体",
                PaymentParty = HainanStage2PaymentParties.Qingneng
            };

            var evaluation = Stage2PreflightPolicy.Evaluate(
                new[] { requirement },
                new IStage2PaymentPartyDecision[] { decision });

            Assert.IsTrue(evaluation.HasOutstandingRequiredDecisions);
            Assert.IsTrue(evaluation.HasInvalidDecisions);
        }

        [TestMethod]
        public void UnsupportedOrDuplicateDecisionCannotContinue()
        {
            var requirement = CreatePaymentPartyRequirement("代理费", "主体甲");
            var unsupported = Stage2PreflightPolicy.Evaluate(
                new[] { requirement },
                new IStage2PaymentPartyDecision[]
                {
                    new HainanStage2SummarySubjectDecision
                    {
                        SettlementKind = "代理费",
                        Entity = "主体甲",
                        PaymentParty = "未知"
                    }
                });
            var duplicate = Stage2PreflightPolicy.Evaluate(
                new[] { requirement },
                new IStage2PaymentPartyDecision[]
                {
                    Decision("主体甲", HainanStage2PaymentParties.Qingneng),
                    Decision("主体甲", HainanStage2PaymentParties.Qinghui)
                });

            Assert.IsTrue(unsupported.HasInvalidDecisions);
            Assert.IsFalse(unsupported.CanContinue);
            Assert.IsTrue(duplicate.HasInvalidDecisions);
            Assert.IsFalse(duplicate.CanContinue);
        }

        [TestMethod]
        public void TemplateRequirementIsOutstandingUntilMatchingDecisionExists()
        {
            var requirement = CreateTemplateRequirement("代理费", "新增主体", "C:\\templates\\a.xlsx", "C:\\templates\\b.xlsx");

            var outstanding = Stage2PreflightPolicy.Evaluate(
                new[] { requirement },
                new IStage2PaymentPartyDecision[0],
                new IStage2TemplateDecision[0]);
            var resolved = Stage2PreflightPolicy.Evaluate(
                new[] { requirement },
                new IStage2PaymentPartyDecision[0],
                new IStage2TemplateDecision[]
                {
                    TemplateDecision("新增主体", "c:\\TEMPLATES\\b.xlsx")
                });

            Assert.IsTrue(outstanding.HasOutstandingRequiredDecisions);
            Assert.IsFalse(outstanding.CanContinue);
            Assert.AreEqual(Stage2TemplateDecisionStatus.Outstanding, outstanding.TemplateDecisionResolutions[0].Status);
            Assert.IsFalse(resolved.HasOutstandingRequiredDecisions);
            Assert.IsTrue(resolved.CanContinue);
            Assert.AreEqual(Stage2TemplateDecisionStatus.Resolved, resolved.TemplateDecisionResolutions[0].Status);
        }

        [TestMethod]
        public void PaymentAndTemplateRequirementsForSameSubjectResolveIndependently()
        {
            var payment = CreatePaymentPartyRequirement("代理费", "新增主体");
            var template = CreateTemplateRequirement("代理费", "新增主体", "C:\\templates\\a.xlsx");

            var evaluation = Stage2PreflightPolicy.Evaluate(
                new[] { payment, template },
                new IStage2PaymentPartyDecision[]
                {
                    Decision("新增主体", HainanStage2PaymentParties.Qingneng)
                },
                new IStage2TemplateDecision[]
                {
                    TemplateDecision("新增主体", "C:\\templates\\a.xlsx")
                });

            Assert.IsTrue(evaluation.CanContinue);
            Assert.AreEqual(Stage2PaymentPartyDecisionStatus.Resolved, evaluation.DecisionResolutions[0].Status);
            Assert.AreEqual(Stage2TemplateDecisionStatus.Resolved, evaluation.TemplateDecisionResolutions[0].Status);
        }

        [TestMethod]
        public void UnsupportedDuplicateOrStaleTemplateDecisionCannotContinue()
        {
            var requirement = CreateTemplateRequirement("代理费", "主体甲", "C:\\templates\\a.xlsx");
            var unsupported = Stage2PreflightPolicy.Evaluate(
                new[] { requirement },
                new IStage2PaymentPartyDecision[0],
                new IStage2TemplateDecision[]
                {
                    TemplateDecision("主体甲", "C:\\templates\\unknown.xlsx")
                });
            var duplicate = Stage2PreflightPolicy.Evaluate(
                new[] { requirement },
                new IStage2PaymentPartyDecision[0],
                new IStage2TemplateDecision[]
                {
                    TemplateDecision("主体甲", "C:\\templates\\a.xlsx"),
                    TemplateDecision("主体甲", "C:\\templates\\a.xlsx")
                });
            var stale = Stage2PreflightPolicy.Evaluate(
                new IStage2PreflightIssue[0],
                new IStage2PaymentPartyDecision[0],
                new IStage2TemplateDecision[]
                {
                    TemplateDecision("主体甲", "C:\\templates\\a.xlsx")
                });

            Assert.IsTrue(unsupported.HasInvalidDecisions);
            Assert.AreEqual(Stage2TemplateDecisionStatus.Invalid, unsupported.TemplateDecisionResolutions[0].Status);
            Assert.IsTrue(duplicate.HasInvalidDecisions);
            Assert.AreEqual(Stage2TemplateDecisionStatus.Conflicting, duplicate.TemplateDecisionResolutions[0].Status);
            Assert.IsTrue(stale.HasInvalidDecisions);
            Assert.AreEqual(Stage2TemplateDecisionStatus.Stale, stale.TemplateDecisionResolutions[0].Status);
        }

        [TestMethod]
        public void OneIssueCannotRequireBothPaymentPartyAndTemplateSelection()
        {
            var issue = CreateTemplateRequirement("代理费", "主体甲", "C:\\templates\\a.xlsx");
            issue.RequiresPaymentPartySelection = true;
            issue.AvailablePaymentParties.AddRange(HainanStage2PaymentParties.Supported);

            var evaluation = Stage2PreflightPolicy.Evaluate(
                new[] { issue },
                new IStage2PaymentPartyDecision[0],
                new IStage2TemplateDecision[0]);

            Assert.IsTrue(evaluation.HasBlockingIssues);
            Assert.IsFalse(evaluation.CanContinue);
            StringAssert.Contains(evaluation.InvalidDefinitions[0], "不能同时要求支付方和模板选择");
        }

        [TestMethod]
        public void MissingStableCodeIsInvalidDefinition()
        {
            var issue = CreateIssue(Stage2PreflightDisposition.Review);
            issue.Code = null;

            var evaluation = Stage2PreflightPolicy.Evaluate(
                new[] { issue },
                new IStage2PaymentPartyDecision[0]);

            Assert.IsTrue(evaluation.HasBlockingIssues);
            Assert.IsFalse(evaluation.CanContinue);
        }

        [TestMethod]
        public void ReportSummaryIgnoresNullIssueEntries()
        {
            var report = new HainanStage2PreflightReport();
            report.Issues.Add(null);
            report.Issues.Add(CreateIssue(Stage2PreflightDisposition.Information));

            Assert.IsFalse(report.HasBlockingIssues);
            Assert.IsFalse(report.HasRequiredDecisions);
            Assert.IsTrue(report.CanGenerate);
        }

        private static HainanStage2CheckIssue CreateIssue(Stage2PreflightDisposition disposition)
        {
            return new HainanStage2CheckIssue
            {
                Code = Stage2PreflightIssueKinds.NewCustomer,
                Disposition = disposition,
                SettlementKind = "代理费",
                Entity = "主体甲"
            };
        }

        private static HainanStage2CheckIssue CreatePaymentPartyRequirement(string kind, string entity)
        {
            var issue = new HainanStage2CheckIssue
            {
                Code = Stage2PreflightIssueKinds.PaymentPartyRequired,
                Disposition = Stage2PreflightDisposition.RequiredDecision,
                SettlementKind = kind,
                Entity = entity,
                RequiresPaymentPartySelection = true
            };
            issue.AvailablePaymentParties.AddRange(HainanStage2PaymentParties.Supported);
            return issue;
        }

        private static HainanStage2CheckIssue CreateTemplateRequirement(
            string kind,
            string entity,
            params string[] templates)
        {
            var issue = new HainanStage2CheckIssue
            {
                Code = Stage2PreflightIssueKinds.AmbiguousBorrowTemplates,
                Disposition = Stage2PreflightDisposition.RequiredDecision,
                SettlementKind = kind,
                Entity = entity,
                RequiresTemplateSelection = true
            };
            issue.AvailableTemplateFiles.AddRange(templates);
            return issue;
        }

        private static HainanStage2SummarySubjectDecision Decision(string entity, string party)
        {
            return new HainanStage2SummarySubjectDecision
            {
                SettlementKind = "代理费",
                Entity = entity,
                PaymentParty = party
            };
        }

        private static HainanStage2TemplateDecision TemplateDecision(string entity, string path)
        {
            return new HainanStage2TemplateDecision
            {
                SettlementKind = "代理费",
                Entity = entity,
                TemplatePath = path
            };
        }
    }
}
