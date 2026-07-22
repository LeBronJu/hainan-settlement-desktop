using System.Linq;
using HainanSettlementTool.Core.Models;
using HainanSettlementTool.Core.Services;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace HainanSettlementTool.Core.Tests
{
    [TestClass]
    public sealed class Stage2RelationshipParameterValidatorTests
    {
        [TestMethod]
        public void BlankSubjectAndBlankParametersMeansNoRelationship()
        {
            var result = Validate(null, Blank(), Blank(), Blank());

            Assert.IsFalse(result.HasRelationship);
            Assert.IsTrue(result.IsValid);
        }

        [TestMethod]
        public void ParameterContentWithoutSubjectIsInvalidIncludingExplicitZero()
        {
            var result = Validate(null, Numeric(0), Blank(), Blank());

            Assert.IsFalse(result.IsValid);
            Assert.AreEqual(Stage2RelationshipParameterErrorKind.ParametersWithoutSubject, result.Errors.Single().Kind);
        }

        [TestMethod]
        public void SubjectRequiresEveryParameterToBeNumericAndPositive()
        {
            var result = Validate("主体", Blank(), Text("百分之五"), Numeric(0));

            CollectionAssert.AreEquivalent(
                new[]
                {
                    Stage2RelationshipParameterErrorKind.MissingParameter,
                    Stage2RelationshipParameterErrorKind.NonNumericParameter,
                    Stage2RelationshipParameterErrorKind.NonPositiveParameter
                },
                result.Errors.Select(item => item.Kind).ToArray());
        }

        [TestMethod]
        public void SubjectWithThreePositiveNumbersIsValid()
        {
            var result = Validate("主体", Numeric(0.5), Numeric(0.8), Numeric(0.13));

            Assert.IsTrue(result.HasRelationship);
            Assert.IsTrue(result.IsValid);
        }

        [TestMethod]
        public void SubjectRejectsNonFiniteNumericParameters()
        {
            var nan = Validate("主体", Numeric(double.NaN), Numeric(0.8), Numeric(0.13));
            var positiveInfinity = Validate("主体", Numeric(0.5), Numeric(double.PositiveInfinity), Numeric(0.13));
            var negativeInfinity = Validate("主体", Numeric(0.5), Numeric(0.8), Numeric(double.NegativeInfinity));

            Assert.AreEqual(Stage2RelationshipParameterErrorKind.NonPositiveParameter, nan.Errors.Single().Kind);
            Assert.AreEqual(Stage2RelationshipParameterErrorKind.NonPositiveParameter, positiveInfinity.Errors.Single().Kind);
            Assert.AreEqual(Stage2RelationshipParameterErrorKind.NonPositiveParameter, negativeInfinity.Errors.Single().Kind);
        }

        private static Stage2RelationshipValidationResult Validate(
            string subject,
            Stage2RelationshipParameterValue ratio,
            Stage2RelationshipParameterValue unitPrice,
            Stage2RelationshipParameterValue taxRate)
        {
            return Stage2RelationshipParameterValidator.Validate(subject, ratio, unitPrice, taxRate);
        }

        private static Stage2RelationshipParameterValue Blank()
        {
            return new Stage2RelationshipParameterValue();
        }

        private static Stage2RelationshipParameterValue Numeric(double value)
        {
            return new Stage2RelationshipParameterValue
            {
                HasContent = true,
                IsNumeric = true,
                Value = value,
                DisplayValue = value.ToString()
            };
        }

        private static Stage2RelationshipParameterValue Text(string value)
        {
            return new Stage2RelationshipParameterValue
            {
                HasContent = true,
                IsNumeric = false,
                DisplayValue = value
            };
        }
    }
}
