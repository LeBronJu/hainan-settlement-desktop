using HainanSettlementTool.Core.Models;
using HainanSettlementTool.Core.Services;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace HainanSettlementTool.Core.Tests
{
    [TestClass]
    public sealed class Stage2PreflightSignatureTests
    {
        [TestMethod]
        public void SignatureIsStableAcrossIssueOrder()
        {
            var first = Issue(Stage2PreflightIssueKinds.NewCustomer, "主体甲", "客户甲");
            var second = Issue(Stage2PreflightIssueKinds.MultipleOwners, "主体乙", "客户乙");

            var left = Stage2PreflightSignature.Create(5, 2, new[] { first, second });
            var right = Stage2PreflightSignature.Create(5, 2, new[] { second, first });

            Assert.IsTrue(Stage2PreflightSignature.Matches(left, right));
        }

        [TestMethod]
        public void SignatureChangesWhenSubjectIssueOrValueChanges()
        {
            var issue = Issue(Stage2PreflightIssueKinds.NewCustomer, "主体甲", "客户甲");
            var baseline = Stage2PreflightSignature.Create(5, 1, new[] { issue });

            issue.CurrentValue = "变更后";
            var changedValue = Stage2PreflightSignature.Create(5, 1, new[] { issue });
            var changedCount = Stage2PreflightSignature.Create(5, 2, new[] { issue });
            var changedMonth = Stage2PreflightSignature.Create(6, 1, new[] { issue });

            Assert.IsFalse(Stage2PreflightSignature.Matches(baseline, changedValue));
            Assert.IsFalse(Stage2PreflightSignature.Matches(baseline, changedCount));
            Assert.IsFalse(Stage2PreflightSignature.Matches(baseline, changedMonth));
            Assert.IsFalse(Stage2PreflightSignature.Matches(null, baseline));
        }

        [TestMethod]
        public void TemplateOptionsAreSignedIndependentOfCandidateOrder()
        {
            var leftIssue = Issue(Stage2PreflightIssueKinds.AmbiguousBorrowTemplates, "新增主体", null);
            leftIssue.Disposition = Stage2PreflightDisposition.RequiredDecision;
            leftIssue.RequiresTemplateSelection = true;
            leftIssue.AvailableTemplateFiles.Add("C:\\templates\\b.xlsx");
            leftIssue.AvailableTemplateFiles.Add("C:\\templates\\a.xlsx");
            var rightIssue = Issue(Stage2PreflightIssueKinds.AmbiguousBorrowTemplates, "新增主体", null);
            rightIssue.Disposition = Stage2PreflightDisposition.RequiredDecision;
            rightIssue.RequiresTemplateSelection = true;
            rightIssue.AvailableTemplateFiles.Add("C:\\templates\\a.xlsx");
            rightIssue.AvailableTemplateFiles.Add("C:\\templates\\b.xlsx");

            var left = Stage2PreflightSignature.Create(5, 1, new[] { leftIssue });
            var right = Stage2PreflightSignature.Create(5, 1, new[] { rightIssue });

            Assert.IsTrue(Stage2PreflightSignature.Matches(left, right));

            rightIssue.AvailableTemplateFiles.Add("C:\\templates\\c.xlsx");
            var changedCandidates = Stage2PreflightSignature.Create(5, 1, new[] { rightIssue });
            rightIssue.AvailableTemplateFiles.RemoveAt(rightIssue.AvailableTemplateFiles.Count - 1);
            rightIssue.RequiresTemplateSelection = false;
            var changedRequirement = Stage2PreflightSignature.Create(5, 1, new[] { rightIssue });

            Assert.IsFalse(Stage2PreflightSignature.Matches(left, changedCandidates));
            Assert.IsFalse(Stage2PreflightSignature.Matches(left, changedRequirement));
        }

        private static HainanStage2CheckIssue Issue(string code, string entity, string customer)
        {
            return new HainanStage2CheckIssue
            {
                Code = code,
                Disposition = Stage2PreflightDisposition.Review,
                SettlementKind = "代理费",
                Entity = entity,
                Customer = customer,
                CurrentValue = "当前值"
            };
        }
    }
}
