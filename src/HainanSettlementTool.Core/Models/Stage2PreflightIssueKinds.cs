namespace HainanSettlementTool.Core.Models
{
    public static class Stage2PreflightIssueKinds
    {
        public const string RelationshipParametersWithoutSubject = "RelationshipParametersWithoutSubject";
        public const string RelationshipParametersInvalid = "RelationshipParametersInvalid";
        public const string FirstOwnerMissing = "FirstOwnerMissing";
        public const string MultipleOwners = "MultipleOwners";
        public const string ConflictingTaxRates = "ConflictingTaxRates";
        public const string DuplicateSummarySubject = "DuplicateSummarySubject";
        public const string ConflictingSummarySources = "ConflictingSummarySources";
        public const string SummarySheetAmbiguous = "SummarySheetAmbiguous";
        public const string SummarySubjectKindInvalid = "SummarySubjectKindInvalid";
        public const string SummaryPaymentSheetMissing = "SummaryPaymentSheetMissing";
        public const string SummaryOrphanSubject = "SummaryOrphanSubject";
        public const string SummaryTargetMonthAlreadyExists = "SummaryTargetMonthAlreadyExists";
        public const string ConflictingPayees = "ConflictingPayees";
        public const string PayeeSourceMissing = "PayeeSourceMissing";
        public const string ConflictingPaymentParties = "ConflictingPaymentParties";
        public const string PaymentPartyRequired = "PaymentPartyRequired";
        public const string NewSummarySubject = "NewSummarySubject";
        public const string NewCustomer = "NewCustomer";
        public const string TemplateMissing = "TemplateMissing";
        public const string TemplateUnreadable = "TemplateUnreadable";
        public const string DuplicateExactTemplates = "DuplicateExactTemplates";
        public const string AmbiguousBorrowTemplates = "AmbiguousBorrowTemplates";
        public const string BorrowedTemplate = "BorrowedTemplate";
        public const string PreviousTemplateCustomerMissing = "PreviousTemplateCustomerMissing";
        public const string RelationshipValueChanged = "RelationshipValueChanged";
        public const string LedgerAmountDifference = "LedgerAmountDifference";
        public const string UnexpectedTargetMonthWorkbook = "UnexpectedTargetMonthWorkbook";
        public const string PlannedTargetMonthWorkbook = "PlannedTargetMonthWorkbook";
        public const string ManagedOutputUnreadable = "ManagedOutputUnreadable";
        public const string PlannedOutputPathConflict = "PlannedOutputPathConflict";
    }
}
