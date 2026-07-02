using System;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using HainanSettlementTool.Core.Models;
using HainanSettlementTool.Core.Services;
using HainanSettlementTool.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace HainanSettlementTool.Excel.Tests
{
    [TestClass]
    public sealed class EmployeeRewardGeneratorTests
    {
        [TestMethod]
        public void ReadLedgerRowsUsesOfficialSheetAndMonthTotalHeaders()
        {
            var root = CreateTempRoot();
            var ledgerPath = Path.Combine(root, "ledger.xlsx");
            var outputDirectory = Path.Combine(root, "out");

            try
            {
                WriteRewardLedger(ledgerPath);

                var rows = new ClosedXmlStage1ExcelGateway().ReadLedgerRows(
                    new EmployeeRewardOptions
                    {
                        Year = 2026,
                        StartMonth = 1,
                        EndMonth = 2,
                        LedgerPath = ledgerPath,
                        OutputDirectory = outputDirectory
                    });

                Assert.AreEqual(3, rows.Count);
                Assert.AreEqual("客户A", rows[0].CustomerName);
                Assert.AreEqual("员工A", rows[0].Owner);
                Assert.AreEqual(10.1234, rows[0].MonthPowers[1], 0.00001);
                Assert.AreEqual(20.5, rows[0].MonthPowers[2], 0.00001);
                Assert.AreEqual("客户B", rows[1].CustomerName);
                Assert.AreEqual("客户C", rows[2].CustomerName);
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void RunGeneratesSummaryPersonalWorkbooksAndReportWithoutTemplates()
        {
            var root = CreateTempRoot();
            var ledgerPath = Path.Combine(root, "ledger.xlsx");
            var outputDirectory = Path.Combine(root, "out");

            try
            {
                WriteRewardLedger(ledgerPath);
                Directory.CreateDirectory(outputDirectory);
                File.WriteAllText(Path.Combine(outputDirectory, "2026年1-2月员工电量奖励-海南.xlsx"), "existing");

                var result = new EmployeeRewardService(new ClosedXmlStage1ExcelGateway()).Run(
                    new EmployeeRewardOptions
                    {
                        Year = 2026,
                        StartMonth = 1,
                        EndMonth = 2,
                        LedgerPath = ledgerPath,
                        OutputDirectory = outputDirectory
                    },
                    null);

                Assert.IsTrue(File.Exists(result.SummaryPath));
                Assert.IsTrue(File.Exists(result.ReportPath));
                Assert.AreEqual(2, result.PersonalWorkbookPaths.Count);
                Assert.AreNotEqual(
                    Path.Combine(outputDirectory, "2026年1-2月员工电量奖励-海南.xlsx"),
                    result.SummaryPath);

                using (var workbook = new XLWorkbook(result.SummaryPath))
                {
                    var detail = workbook.Worksheet("1-2月企业用电量明细");
                    Assert.AreEqual("2026年1-2月企业用电量明细", detail.Cell("A1").GetString());
                    Assert.AreEqual("1月总实际电量（万千瓦时）", detail.Cell(2, 8).GetString());
                    Assert.AreEqual("2月总实际电量（万千瓦时）", detail.Cell(2, 9).GetString());
                    Assert.AreEqual("电量合计（万千瓦时）", detail.Cell(2, 10).GetString());
                    Assert.AreEqual("备注", detail.Cell(2, 11).GetString());
                    Assert.AreEqual("SUM(H3:I3)", detail.Cell(3, 10).FormulaA1);
                    Assert.AreEqual("合计", detail.Cell(6, 1).GetString());
                    Assert.AreEqual("SUM(H3:H5)", detail.Cell(6, 8).FormulaA1);
                    AssertRangeHasBorders(detail, 1, 1, 6, 11);
                    AssertNoBorder(detail.Cell(7, 10));

                    var summary = workbook.Worksheet("1月-2月员工电量汇总");
                    Assert.AreEqual("2026年1-2月员工电量奖励", summary.Cell("A1").GetString());
                    Assert.AreEqual("员工A", summary.Cell(3, 2).GetString());
                    Assert.AreEqual("SUM(C3:D3)", summary.Cell(3, 5).FormulaA1);
                    Assert.AreEqual("E3*10000*0.0001", summary.Cell(3, 6).FormulaA1);
                    Assert.AreEqual("合计", summary.Cell(5, 1).GetString());
                    Assert.IsTrue(summary.MergedRanges.Any(range => range.RangeAddress.FirstAddress.RowNumber == 5
                        && range.RangeAddress.FirstAddress.ColumnNumber == 1
                        && range.RangeAddress.LastAddress.RowNumber == 5
                        && range.RangeAddress.LastAddress.ColumnNumber == 2));
                    Assert.AreEqual("SUM(C3:C4)", summary.Cell(5, 3).FormulaA1);
                    Assert.AreEqual("SUM(D3:D4)", summary.Cell(5, 4).FormulaA1);
                    Assert.AreEqual("SUM(E3:E4)", summary.Cell(5, 5).FormulaA1);
                    Assert.AreEqual("SUM(F3:F4)", summary.Cell(5, 6).FormulaA1);
                    Assert.AreEqual("备注：按0.0001元/千瓦时奖励。", summary.Cell(6, 1).GetString());
                    AssertRangeHasBorders(summary, 1, 1, 6, 7);
                    AssertNoBorder(summary.Cell(7, 1));
                }

                var employeeAPath = result.PersonalWorkbookPaths.Single(path => Path.GetFileName(path).StartsWith("员工A-", StringComparison.Ordinal));
                using (var workbook = new XLWorkbook(employeeAPath))
                {
                    var sheet = workbook.Worksheet("1-2月企业用电明细");
                    Assert.AreEqual("2026年1-2月员工电量确认表(员工A)", sheet.Cell("A1").GetString());
                    Assert.AreEqual("客户A", sheet.Cell(3, 3).GetString());
                    Assert.AreEqual("客户B", sheet.Cell(4, 3).GetString());
                    Assert.AreEqual("合计", sheet.Cell(5, 1).GetString());
                    Assert.AreEqual("SUM(H3:H4)", sheet.Cell(5, 8).FormulaA1);
                    AssertRangeHasBorders(sheet, 1, 1, 5, 11);
                    AssertNoBorder(sheet.Cell(6, 9));
                }
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        private static void WriteRewardLedger(string path)
        {
            using (var workbook = new XLWorkbook())
            {
                workbook.AddWorksheet("废弃").Cell("A1").Value = "不是正式台账";
                var worksheet = workbook.AddWorksheet(LedgerLayout.MainSheetName);
                worksheet.Cell(1, 1).Value = "海南售电结算台账";
                worksheet.Cell(2, 2).Value = "用电企业编号";
                worksheet.Cell(2, 3).Value = "用电企业名称";
                worksheet.Cell(2, 6).Value = "履约开始月份";
                worksheet.Cell(2, 8).Value = "项目开发人";
                worksheet.Cell(2, 9).Value = "代理或自营";
                worksheet.Cell(2, 10).Value = "负责人";

                WriteMonthHeader(worksheet, 1);
                WriteMonthHeader(worksheet, 2);

                WriteLedgerRow(worksheet, 4, 1, "001", "客户A", "202601", "代理A", "代理", "员工A", 10.1234, 20.5);
                WriteLedgerRow(worksheet, 5, 2, "002", "客户B", "202602", "代理B", "代理", "员工A", 0, 30);
                WriteLedgerRow(worksheet, 6, 3, "003", "客户C", "202603", "自营", "自营", "员工B", 5, 0);

                var jan = LedgerLayout.MonthStartColumn(1);
                var feb = LedgerLayout.MonthStartColumn(2);
                worksheet.Cell(7, jan).Value = 15.1234;
                worksheet.Cell(7, feb).Value = 50.5;
                worksheet.Column(jan).Hide();
                worksheet.Column(feb).Hide();

                Directory.CreateDirectory(Path.GetDirectoryName(path));
                workbook.SaveAs(path);
            }
        }

        private static void WriteMonthHeader(IXLWorksheet worksheet, int month)
        {
            var column = LedgerLayout.MonthStartColumn(month);
            worksheet.Cell(1, column).Value = month + "月";
            worksheet.Cell(2, column).Value = "总实际电量（万千瓦时）";
        }

        private static void WriteLedgerRow(
            IXLWorksheet worksheet,
            int row,
            int sequence,
            string customerCode,
            string customerName,
            string contractStartMonth,
            string developer,
            string agentType,
            string owner,
            double janPower,
            double febPower)
        {
            worksheet.Cell(row, 1).Value = sequence;
            worksheet.Cell(row, 2).Value = customerCode;
            worksheet.Cell(row, 3).Value = customerName;
            worksheet.Cell(row, 6).Value = contractStartMonth;
            worksheet.Cell(row, 8).Value = developer;
            worksheet.Cell(row, 9).Value = agentType;
            worksheet.Cell(row, 10).Value = owner;
            worksheet.Cell(row, LedgerLayout.MonthStartColumn(1)).Value = janPower;
            worksheet.Cell(row, LedgerLayout.MonthStartColumn(2)).Value = febPower;
        }

        private static void AssertRangeHasBorders(IXLWorksheet worksheet, int firstRow, int firstColumn, int lastRow, int lastColumn)
        {
            for (var row = firstRow; row <= lastRow; row++)
            {
                for (var column = firstColumn; column <= lastColumn; column++)
                {
                    AssertHasBorder(worksheet.Cell(row, column), "row " + row + ", column " + column);
                }
            }
        }

        private static void AssertHasBorder(IXLCell cell, string location)
        {
            Assert.AreNotEqual(XLBorderStyleValues.None, cell.Style.Border.TopBorder, location + " top border");
            Assert.AreNotEqual(XLBorderStyleValues.None, cell.Style.Border.BottomBorder, location + " bottom border");
            Assert.AreNotEqual(XLBorderStyleValues.None, cell.Style.Border.LeftBorder, location + " left border");
            Assert.AreNotEqual(XLBorderStyleValues.None, cell.Style.Border.RightBorder, location + " right border");
            Assert.AreEqual(XLColor.Black, cell.Style.Border.TopBorderColor, location + " top border color");
            Assert.AreEqual(XLColor.Black, cell.Style.Border.BottomBorderColor, location + " bottom border color");
            Assert.AreEqual(XLColor.Black, cell.Style.Border.LeftBorderColor, location + " left border color");
            Assert.AreEqual(XLColor.Black, cell.Style.Border.RightBorderColor, location + " right border color");
        }

        private static void AssertNoBorder(IXLCell cell)
        {
            Assert.AreEqual(XLBorderStyleValues.None, cell.Style.Border.TopBorder);
            Assert.AreEqual(XLBorderStyleValues.None, cell.Style.Border.BottomBorder);
            Assert.AreEqual(XLBorderStyleValues.None, cell.Style.Border.LeftBorder);
            Assert.AreEqual(XLBorderStyleValues.None, cell.Style.Border.RightBorder);
        }

        private static string CreateTempRoot()
        {
            return Path.Combine(Path.GetTempPath(), "HainanSettlementToolTests", Guid.NewGuid().ToString("N"));
        }

        private static void DeleteTempRoot(string root)
        {
            if (Directory.Exists(root))
            {
                Directory.Delete(root, true);
            }
        }
    }
}
