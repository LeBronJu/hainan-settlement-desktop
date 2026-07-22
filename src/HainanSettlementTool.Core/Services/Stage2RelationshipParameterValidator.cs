using System;
using System.Collections.Generic;
using System.Linq;
using HainanSettlementTool.Core.Models;

namespace HainanSettlementTool.Core.Services
{
    public static class Stage2RelationshipParameterValidator
    {
        public static Stage2RelationshipValidationResult Validate(
            string subject,
            Stage2RelationshipParameterValue ratio,
            Stage2RelationshipParameterValue unitPrice,
            Stage2RelationshipParameterValue taxRate)
        {
            var result = new Stage2RelationshipValidationResult
            {
                HasRelationship = !string.IsNullOrWhiteSpace(subject)
            };
            var parameters = new[]
            {
                Pair("比例", ratio),
                Pair("单价", unitPrice),
                Pair("税率", taxRate)
            };

            if (!result.HasRelationship)
            {
                foreach (var parameter in parameters.Where(item => item.Value.HasContent))
                {
                    result.Errors.Add(new Stage2RelationshipParameterError
                    {
                        Kind = Stage2RelationshipParameterErrorKind.ParametersWithoutSubject,
                        ParameterName = parameter.Name,
                        DisplayValue = parameter.Value.DisplayValue
                    });
                }

                return result;
            }

            foreach (var parameter in parameters)
            {
                if (!parameter.Value.HasContent)
                {
                    result.Errors.Add(Error(Stage2RelationshipParameterErrorKind.MissingParameter, parameter));
                }
                else if (!parameter.Value.IsNumeric)
                {
                    result.Errors.Add(Error(Stage2RelationshipParameterErrorKind.NonNumericParameter, parameter));
                }
                else if (double.IsNaN(parameter.Value.Value)
                    || double.IsInfinity(parameter.Value.Value)
                    || parameter.Value.Value <= 0)
                {
                    result.Errors.Add(Error(Stage2RelationshipParameterErrorKind.NonPositiveParameter, parameter));
                }
            }

            return result;
        }

        private static ParameterPair Pair(string name, Stage2RelationshipParameterValue value)
        {
            return new ParameterPair
            {
                Name = name,
                Value = value ?? new Stage2RelationshipParameterValue()
            };
        }

        private static Stage2RelationshipParameterError Error(
            Stage2RelationshipParameterErrorKind kind,
            ParameterPair parameter)
        {
            return new Stage2RelationshipParameterError
            {
                Kind = kind,
                ParameterName = parameter.Name,
                DisplayValue = parameter.Value.DisplayValue
            };
        }

        private sealed class ParameterPair
        {
            public string Name { get; set; }
            public Stage2RelationshipParameterValue Value { get; set; }
        }
    }
}
