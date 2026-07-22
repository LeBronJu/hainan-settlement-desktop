using System.Collections.Generic;

namespace HainanSettlementTool.Core.Models
{
    public interface IStage2PreflightIssue
    {
        string Code { get; }
        Stage2PreflightDisposition Disposition { get; }
        string SettlementKind { get; }
        string Entity { get; }
        bool RequiresPaymentPartySelection { get; }
        IReadOnlyList<string> PaymentPartyOptions { get; }
    }
}
