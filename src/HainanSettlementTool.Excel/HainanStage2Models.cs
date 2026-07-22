using System.Collections.Generic;
using System.Linq;
using HainanSettlementTool.Core.Models;

namespace HainanSettlementTool.Excel
{
    internal sealed class HainanStage2LedgerSnapshot
    {
        public List<HainanStage2DetailSettlementRow> ProxyRows { get; } = new List<HainanStage2DetailSettlementRow>();
        public List<HainanStage2DetailSettlementRow> IntermediaryRows { get; } = new List<HainanStage2DetailSettlementRow>();
        public List<HainanStage2RelationshipOccurrence> Relationships { get; } = new List<HainanStage2RelationshipOccurrence>();
        public List<HainanStage2SubjectGroup> SubjectGroups { get; } = new List<HainanStage2SubjectGroup>();
        public List<HainanStage2CheckIssue> Issues { get; } = new List<HainanStage2CheckIssue>();
    }

    internal sealed class HainanStage2RelationshipOccurrence
    {
        public int LedgerRow { get; set; }
        public string Customer { get; set; }
        public string Owner { get; set; }
        public string Entity { get; set; }
        public string Kind { get; set; }
        public string SettlementKind { get; set; }
        public double TaxRate { get; set; }
    }

    internal sealed class HainanStage2SubjectGroup
    {
        public string Kind { get; set; }
        public string SettlementKind { get; set; }
        public string Entity { get; set; }
        public string Owner { get; set; }
        public int FirstLedgerRow { get; set; }
        public double TaxRate { get; set; }
        public List<string> Owners { get; } = new List<string>();
        public List<HainanStage2DetailSettlementRow> Rows { get; } = new List<HainanStage2DetailSettlementRow>();
    }

    internal sealed class HainanStage2TemplateCandidate
    {
        public string Kind { get; set; }
        public string Entity { get; set; }
        public string Owner { get; set; }
        public string Path { get; set; }
    }

    internal sealed class HainanStage2TemplateCatalog
    {
        public List<HainanStage2TemplateCandidate> Candidates { get; } = new List<HainanStage2TemplateCandidate>();
        public List<HainanStage2CheckIssue> Issues { get; } = new List<HainanStage2CheckIssue>();

        public List<HainanStage2TemplateCandidate> ExactCandidates(string kind, string entity)
        {
            var key = HainanStage2ExcelUtil.TemplateSubjectKey(kind, entity);
            return Candidates
                .Where(candidate => HainanStage2ExcelUtil.TemplateSubjectKey(candidate.Kind, candidate.Entity) == key)
                .OrderBy(candidate => candidate.Path, System.StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        public List<HainanStage2TemplateCandidate> CandidatesForKind(string kind)
        {
            return Candidates
                .Where(candidate => candidate.Kind == kind)
                .OrderBy(candidate => candidate.Path, System.StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        public HainanStage2TemplateCandidate UniqueForKind(string kind)
        {
            var candidates = CandidatesForKind(kind);
            return candidates.Count == 1 ? candidates[0] : null;
        }
    }

    internal sealed class HainanStage2SummaryMetaRow
    {
        public int Row { get; set; }
        public string SheetName { get; set; }
        public string Entity { get; set; }
        public string Kind { get; set; }
        public string Payee { get; set; }
        public string PaymentParty { get; set; }
    }
}
