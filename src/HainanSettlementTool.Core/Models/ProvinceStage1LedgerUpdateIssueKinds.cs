namespace HainanSettlementTool.Core.Models
{
    public static class ProvinceStage1LedgerUpdateIssueKinds
    {
        public const string MonthMismatch = "MonthMismatch";
        public const string MultiAccountCustomer = "MultiAccountCustomer";
        public const string ManualMatchedCustomer = "ManualMatchedCustomer";
        public const string CreatedCustomer = "CreatedCustomer";
        public const string SkippedPowerCustomer = "SkippedPowerCustomer";
        public const string ExistingPowerDifference = "ExistingPowerDifference";
        public const string PowerCustomerMissingInLedger = "PowerCustomerMissingInLedger";
        public const string LedgerCustomerMissingInPower = "LedgerCustomerMissingInPower";
        public const string PossibleAlias = "PossibleAlias";
    }
}
