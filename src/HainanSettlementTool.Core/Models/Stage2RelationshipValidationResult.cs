using System.Collections.Generic;

namespace HainanSettlementTool.Core.Models
{
    public enum Stage2RelationshipParameterErrorKind
    {
        ParametersWithoutSubject = 0,
        MissingParameter = 1,
        NonNumericParameter = 2,
        NonPositiveParameter = 3
    }

    public sealed class Stage2RelationshipParameterError
    {
        public Stage2RelationshipParameterErrorKind Kind { get; set; }
        public string ParameterName { get; set; }
        public string DisplayValue { get; set; }
    }

    public sealed class Stage2RelationshipValidationResult
    {
        public bool HasRelationship { get; set; }
        public List<Stage2RelationshipParameterError> Errors { get; } = new List<Stage2RelationshipParameterError>();

        public bool IsValid
        {
            get { return Errors.Count == 0; }
        }
    }
}
