using System.Collections.Generic;

namespace HainanSettlementTool.Core.Models
{
    public abstract class Stage2PreflightIssue : IStage2PreflightIssue
    {
        public string Code { get; set; }
        public Stage2PreflightDisposition Disposition { get; set; }
        public string Severity { get; set; }
        public string Category { get; set; }
        public string Kind { get; set; }
        public string SettlementKind { get; set; }
        public string Customer { get; set; }
        public string Owner { get; set; }
        public string Entity { get; set; }
        public int LedgerRow { get; set; }
        public string TemplateFile { get; set; }
        public string SheetName { get; set; }
        public string PreviousValue { get; set; }
        public string CurrentValue { get; set; }
        public string Message { get; set; }
        public string Suggestion { get; set; }
        public bool RequiresPaymentPartySelection { get; set; }
        public List<string> AvailablePaymentParties { get; } = new List<string>();

        public IReadOnlyList<string> PaymentPartyOptions
        {
            get { return AvailablePaymentParties; }
        }

        public bool BlocksGeneration
        {
            get { return Disposition == Stage2PreflightDisposition.Blocker; }
        }

        public bool RequiresDecision
        {
            get
            {
                return Disposition == Stage2PreflightDisposition.RequiredDecision
                    || RequiresPaymentPartySelection;
            }
        }
    }
}
