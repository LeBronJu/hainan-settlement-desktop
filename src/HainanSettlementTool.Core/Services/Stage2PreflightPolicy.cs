using System;
using System.Collections.Generic;
using System.Linq;
using HainanSettlementTool.Core.Models;

namespace HainanSettlementTool.Core.Services
{
    public static class Stage2PreflightPolicy
    {
        public static Stage2PreflightEvaluation Evaluate(
            IEnumerable<IStage2PreflightIssue> issues,
            IEnumerable<IStage2PaymentPartyDecision> decisions)
        {
            return Evaluate(issues, decisions, Enumerable.Empty<IStage2TemplateDecision>());
        }

        public static Stage2PreflightEvaluation Evaluate(
            IEnumerable<IStage2PreflightIssue> issues,
            IEnumerable<IStage2PaymentPartyDecision> decisions,
            IEnumerable<IStage2TemplateDecision> templateDecisions)
        {
            var issueList = (issues ?? Enumerable.Empty<IStage2PreflightIssue>())
                .Where(issue => issue != null)
                .ToList();
            var decisionList = (decisions ?? Enumerable.Empty<IStage2PaymentPartyDecision>())
                .Where(decision => decision != null)
                .ToList();
            var templateDecisionList = (templateDecisions ?? Enumerable.Empty<IStage2TemplateDecision>())
                .Where(decision => decision != null)
                .ToList();
            var evaluation = new Stage2PreflightEvaluation();

            foreach (var issue in issueList)
            {
                if (string.IsNullOrWhiteSpace(issue.Code))
                {
                    evaluation.InvalidDefinitions.Add("阶段二预检项目缺少稳定问题代码。");
                }

                if (issue.Disposition == Stage2PreflightDisposition.Blocker)
                {
                    evaluation.BlockerCount++;
                }
                else if (issue.Disposition == Stage2PreflightDisposition.Review)
                {
                    evaluation.ReviewCount++;
                }
                else if (issue.Disposition == Stage2PreflightDisposition.Information)
                {
                    evaluation.InformationCount++;
                }

                if (issue.RequiresPaymentPartySelection && issue.RequiresTemplateSelection)
                {
                    evaluation.InvalidDefinitions.Add(
                        "同一阶段二预检项目不能同时要求支付方和模板选择：" + issue.Code + "。");
                }
                else if (issue.Disposition == Stage2PreflightDisposition.RequiredDecision
                    && !issue.RequiresPaymentPartySelection
                    && !issue.RequiresTemplateSelection)
                {
                    evaluation.InvalidDefinitions.Add("阶段二必选项目缺少受支持的决策类型：" + issue.Code + "。");
                }
            }

            var requirements = issueList
                .Where(issue => issue.RequiresPaymentPartySelection && !issue.RequiresTemplateSelection)
                .ToList();
            var requirementKeys = new HashSet<string>();
            foreach (var requirement in requirements)
            {
                var key = DecisionKey(requirement.SettlementKind, requirement.Entity);
                if (string.IsNullOrWhiteSpace(requirement.SettlementKind)
                    || string.IsNullOrWhiteSpace(requirement.Entity))
                {
                    evaluation.InvalidDefinitions.Add("支付方必选项目缺少费用类型或主体：" + requirement.Code + "。");
                    continue;
                }

                if (requirement.PaymentPartyOptions == null || requirement.PaymentPartyOptions.Count == 0)
                {
                    evaluation.InvalidDefinitions.Add("支付方必选项目没有可选支付方：" + requirement.SettlementKind + " " + requirement.Entity + "。");
                    continue;
                }

                if (!requirementKeys.Add(key))
                {
                    evaluation.InvalidDefinitions.Add("同一汇总主体存在重复支付方必选项目：" + requirement.SettlementKind + " " + requirement.Entity + "。");
                    continue;
                }

                var matches = decisionList
                    .Where(decision => DecisionKey(decision.SettlementKind, decision.Entity) == key)
                    .ToList();
                evaluation.DecisionResolutions.Add(ResolvePaymentParty(requirement, matches));
            }

            foreach (var staleGroup in decisionList
                .GroupBy(decision => DecisionKey(decision.SettlementKind, decision.Entity))
                .Where(group => !requirementKeys.Contains(group.Key)))
            {
                var first = staleGroup.First();
                evaluation.DecisionResolutions.Add(new Stage2PaymentPartyDecisionResolution
                {
                    SettlementKind = first.SettlementKind,
                    Entity = first.Entity,
                    PaymentParty = first.PaymentParty,
                    Status = Stage2PaymentPartyDecisionStatus.Stale,
                    Message = "支付方选择没有对应的本次预检项目。"
                });
            }

            var templateRequirements = issueList
                .Where(issue => issue.RequiresTemplateSelection && !issue.RequiresPaymentPartySelection)
                .ToList();
            var templateRequirementKeys = new HashSet<string>();
            foreach (var requirement in templateRequirements)
            {
                var key = DecisionKey(requirement.SettlementKind, requirement.Entity);
                if (string.IsNullOrWhiteSpace(requirement.SettlementKind)
                    || string.IsNullOrWhiteSpace(requirement.Entity))
                {
                    evaluation.InvalidDefinitions.Add("模板必选项目缺少费用类型或主体：" + requirement.Code + "。");
                    continue;
                }

                if (requirement.TemplateOptions == null || requirement.TemplateOptions.Count == 0)
                {
                    evaluation.InvalidDefinitions.Add("模板必选项目没有可选模板：" + requirement.SettlementKind + " " + requirement.Entity + "。");
                    continue;
                }

                if (!templateRequirementKeys.Add(key))
                {
                    evaluation.InvalidDefinitions.Add("同一汇总主体存在重复模板必选项目：" + requirement.SettlementKind + " " + requirement.Entity + "。");
                    continue;
                }

                var matches = templateDecisionList
                    .Where(decision => DecisionKey(decision.SettlementKind, decision.Entity) == key)
                    .ToList();
                evaluation.TemplateDecisionResolutions.Add(ResolveTemplate(requirement, matches));
            }

            foreach (var staleGroup in templateDecisionList
                .GroupBy(decision => DecisionKey(decision.SettlementKind, decision.Entity))
                .Where(group => !templateRequirementKeys.Contains(group.Key)))
            {
                var first = staleGroup.First();
                evaluation.TemplateDecisionResolutions.Add(new Stage2TemplateDecisionResolution
                {
                    SettlementKind = first.SettlementKind,
                    Entity = first.Entity,
                    TemplatePath = first.TemplatePath,
                    Status = Stage2TemplateDecisionStatus.Stale,
                    Message = "模板选择没有对应的本次预检项目。"
                });
            }

            return evaluation;
        }

        public static int Count(
            IEnumerable<Stage2PreflightIssue> issues,
            Stage2PreflightDisposition disposition)
        {
            return Safe(issues).Count(issue => issue.Disposition == disposition);
        }

        public static bool HasBlockingIssues(IEnumerable<Stage2PreflightIssue> issues)
        {
            return Safe(issues).Any(issue => issue.BlocksGeneration);
        }

        public static bool HasRequiredDecisions(IEnumerable<Stage2PreflightIssue> issues)
        {
            return Safe(issues).Any(issue => issue.RequiresDecision);
        }

        public static bool CanGenerate(IEnumerable<Stage2PreflightIssue> issues)
        {
            return !HasBlockingIssues(issues) && !HasRequiredDecisions(issues);
        }

        private static Stage2PaymentPartyDecisionResolution ResolvePaymentParty(
            IStage2PreflightIssue requirement,
            IList<IStage2PaymentPartyDecision> matches)
        {
            if (matches.Count == 0)
            {
                return Resolution(requirement, null, Stage2PaymentPartyDecisionStatus.Outstanding, "尚未选择支付方。");
            }

            if (matches.Count > 1)
            {
                return Resolution(requirement, null, Stage2PaymentPartyDecisionStatus.Conflicting, "同一汇总主体存在多个支付方选择。");
            }

            var selected = matches[0].PaymentParty;
            if (string.IsNullOrWhiteSpace(selected)
                || !requirement.PaymentPartyOptions.Contains(selected, StringComparer.Ordinal))
            {
                return Resolution(requirement, selected, Stage2PaymentPartyDecisionStatus.Invalid, "选择的支付方不在可选范围内。");
            }

            return Resolution(requirement, selected, Stage2PaymentPartyDecisionStatus.Resolved, "支付方选择已完成。");
        }

        private static Stage2PaymentPartyDecisionResolution Resolution(
            IStage2PreflightIssue requirement,
            string paymentParty,
            Stage2PaymentPartyDecisionStatus status,
            string message)
        {
            return new Stage2PaymentPartyDecisionResolution
            {
                SettlementKind = requirement.SettlementKind,
                Entity = requirement.Entity,
                PaymentParty = paymentParty,
                Status = status,
                Message = message
            };
        }

        private static Stage2TemplateDecisionResolution ResolveTemplate(
            IStage2PreflightIssue requirement,
            IList<IStage2TemplateDecision> matches)
        {
            if (matches.Count == 0)
            {
                return TemplateResolution(
                    requirement,
                    null,
                    Stage2TemplateDecisionStatus.Outstanding,
                    "尚未选择借用模板。");
            }

            if (matches.Count > 1)
            {
                return TemplateResolution(
                    requirement,
                    null,
                    Stage2TemplateDecisionStatus.Conflicting,
                    "同一汇总主体存在多个模板选择。");
            }

            var selected = matches[0].TemplatePath;
            if (string.IsNullOrWhiteSpace(selected)
                || !requirement.TemplateOptions.Any(option =>
                    string.Equals(TextUtil.S(option), TextUtil.S(selected), StringComparison.OrdinalIgnoreCase)))
            {
                return TemplateResolution(
                    requirement,
                    selected,
                    Stage2TemplateDecisionStatus.Invalid,
                    "选择的模板不在本次预检候选范围内。");
            }

            return TemplateResolution(
                requirement,
                selected,
                Stage2TemplateDecisionStatus.Resolved,
                "借用模板选择已完成。");
        }

        private static Stage2TemplateDecisionResolution TemplateResolution(
            IStage2PreflightIssue requirement,
            string templatePath,
            Stage2TemplateDecisionStatus status,
            string message)
        {
            return new Stage2TemplateDecisionResolution
            {
                SettlementKind = requirement.SettlementKind,
                Entity = requirement.Entity,
                TemplatePath = templatePath,
                Status = status,
                Message = message
            };
        }

        private static string DecisionKey(string settlementKind, string entity)
        {
            return TextUtil.S(settlementKind) + "|" + TextUtil.CustomerKey(entity);
        }

        private static IEnumerable<Stage2PreflightIssue> Safe(IEnumerable<Stage2PreflightIssue> issues)
        {
            return (issues ?? Enumerable.Empty<Stage2PreflightIssue>())
                .Where(issue => issue != null);
        }
    }
}
