using System.Collections.Generic;

namespace HainanSettlementTool.Core.Models
{
    public sealed class Stage2CheckIssue
    {
        public string Severity { get; set; }
        public string Category { get; set; }
        public string Kind { get; set; }
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
    }
}
