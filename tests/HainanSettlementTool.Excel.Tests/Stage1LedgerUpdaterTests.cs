using System;
using System.Linq;
using System.IO;
using System.Text;
using ClosedXML.Excel;
using HainanSettlementTool.Core.Models;
using HainanSettlementTool.Core.Services;
using HainanSettlementTool.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace HainanSettlementTool.Excel.Tests
{
    [TestClass]
    public sealed class Stage1LedgerUpdaterTests
    {
        [TestMethod]
        public void RunInsertsNewCustomerBeforeFooterAndWritesMonthPower()
        {
            var root = CreateTempRoot();
            var baseLedgerPath = Path.Combine(root, "base-ledger.xlsx");
            var powerPath = Path.Combine(root, "power.xlsx");
            var outputDirectory = Path.Combine(root, "out");

            try
            {
                WriteBaseLedger(baseLedgerPath);
                WritePowerWorkbook(powerPath, "新增客户", 12.3456, 1.1, 2.2, 3.3, 4.4);

                var report = RunStage1(baseLedgerPath, powerPath, outputDirectory);

                Assert.AreEqual(1, report.NewRows);
                Assert.AreEqual(1, report.NewCustomers.Count);
                Assert.AreEqual("新增客户", report.NewCustomers[0].Name);
                Assert.AreEqual(5, report.NewCustomers[0].TargetRow);

                using (var workbook = new XLWorkbook(report.Output))
                {
                    var worksheet = workbook.Worksheet(LedgerLayout.MainSheetName);
                    var targetStart = LedgerLayout.MonthStartColumn(4);

                    Assert.AreEqual("既有客户", worksheet.Cell(4, 3).GetString());
                    Assert.AreEqual("新增客户", worksheet.Cell(5, 3).GetString());
                    Assert.AreEqual("合计", worksheet.Cell(6, 1).GetString());
                    Assert.AreEqual(12.3456, worksheet.Cell(5, targetStart).GetDouble(), 0.00001);
                    Assert.AreEqual(1.1, worksheet.Cell(5, targetStart + 1).GetDouble(), 0.00001);
                    Assert.AreEqual(2.2, worksheet.Cell(5, targetStart + 2).GetDouble(), 0.00001);
                    Assert.AreEqual(3.3, worksheet.Cell(5, targetStart + 3).GetDouble(), 0.00001);
                    Assert.AreEqual(4.4, worksheet.Cell(5, targetStart + 4).GetDouble(), 0.00001);
                }
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void RunFillsOnlyUnambiguousCustomerCodesFromRawDetail()
        {
            var root = CreateTempRoot();
            var baseLedgerPath = Path.Combine(root, "base-ledger.xlsx");
            var rawDetailPath = Path.Combine(root, "raw.csv");
            var cleanPowerPath = Path.Combine(root, "out", "clean-power.xlsx");
            var outputDirectory = Path.Combine(root, "out");

            try
            {
                WriteBaseLedger(baseLedgerPath);
                WriteRawDetailCsv(
                    rawDetailPath,
                    RawDetailLine("CODE-ONLY", "唯一客户", 10, 1, 2, 3, 4),
                    RawDetailLine("CODE-A", "冲突客户", 5, 1, 1, 1, 2),
                    RawDetailLine("CODE-B", "冲突客户", 6, 2, 2, 1, 1));

                var report = RunStage1(
                    new Stage1Options
                    {
                        Month = 4,
                        BaseLedgerPath = baseLedgerPath,
                        PowerPath = cleanPowerPath,
                        RawDetailPath = rawDetailPath,
                        OutputDirectory = outputDirectory,
                        OutputLedgerName = "updated-ledger.xlsx"
                    });

                using (var workbook = new XLWorkbook(report.Output))
                {
                    var worksheet = workbook.Worksheet(LedgerLayout.MainSheetName);
                    var uniqueRow = FindCustomerRow(worksheet, "唯一客户");
                    var conflictingRow = FindCustomerRow(worksheet, "冲突客户");

                    Assert.AreEqual("CODE-ONLY", worksheet.Cell(uniqueRow, 2).GetString());
                    Assert.AreEqual(string.Empty, worksheet.Cell(conflictingRow, 2).GetString());
                }

                Assert.AreEqual(1, report.CodeFilledFromRaw.Count);
                Assert.AreEqual("唯一客户", report.CodeFilledFromRaw[0].Name);
                CollectionAssert.Contains(report.MissingCodes, "冲突客户");
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void RunWritesLedgerCopyWithoutChangingBaseLedger()
        {
            var root = CreateTempRoot();
            var baseLedgerPath = Path.Combine(root, "base-ledger.xlsx");
            var powerPath = Path.Combine(root, "power.xlsx");
            var outputDirectory = Path.Combine(root, "out");

            try
            {
                WriteBaseLedger(baseLedgerPath);
                WritePowerWorkbook(powerPath, "新增客户", 12.3456, 1.1, 2.2, 3.3, 4.4);

                var report = RunStage1(baseLedgerPath, powerPath, outputDirectory);

                Assert.AreNotEqual(Path.GetFullPath(baseLedgerPath), Path.GetFullPath(report.Output));
                using (var baseWorkbook = new XLWorkbook(baseLedgerPath))
                using (var outputWorkbook = new XLWorkbook(report.Output))
                {
                    var baseSheet = baseWorkbook.Worksheet(LedgerLayout.MainSheetName);
                    var outputSheet = outputWorkbook.Worksheet(LedgerLayout.MainSheetName);

                    Assert.AreEqual("合计", baseSheet.Cell(5, 1).GetString());
                    Assert.AreEqual("合计", outputSheet.Cell(6, 1).GetString());
                    Assert.AreEqual(-1, TryFindCustomerRow(baseSheet, "新增客户"));
                    Assert.AreNotEqual(-1, TryFindCustomerRow(outputSheet, "新增客户"));
                }
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }


        private static Stage1Report RunStage1(string baseLedgerPath, string powerPath, string outputDirectory)
        {
            return RunStage1(
                new Stage1Options
                {
                    Month = 4,
                    BaseLedgerPath = baseLedgerPath,
                    PowerPath = powerPath,
                    OutputDirectory = outputDirectory,
                    OutputLedgerName = "updated-ledger.xlsx"
                });
        }

        private static Stage1Report RunStage1(Stage1Options options)
        {
            var service = new Stage1Service(new ClosedXmlStage1ExcelGateway());
            return service.Run(options, null);
        }

        private static void WriteBaseLedger(string path)
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.AddWorksheet(LedgerLayout.MainSheetName);
                var marchStart = LedgerLayout.MonthStartColumn(3);

                worksheet.Cell(1, 1).Value = "售电结算台账";
                worksheet.Cell(1, marchStart).Value = "3月";
                worksheet.Cell(4, 1).Value = 1;
                worksheet.Cell(4, 2).Value = "OLD-CODE";
                worksheet.Cell(4, 3).Value = "既有客户";
                worksheet.Cell(4, 10).Value = "负责人";
                worksheet.Cell(5, 1).Value = "合计";
                worksheet.Cell(5, 3).Value = "footer";

                Directory.CreateDirectory(Path.GetDirectoryName(path));
                workbook.SaveAs(path);
            }
        }

        private static void WritePowerWorkbook(
            string path,
            string customer,
            double total,
            double sharp,
            double peak,
            double flat,
            double valley)
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.AddWorksheet("Sheet1");
                worksheet.Cell(1, 1).Value = "零售用户名称";
                worksheet.Cell(1, 2).Value = "总电量(I)";
                worksheet.Cell(1, 3).Value = "尖段电量(L)";
                worksheet.Cell(1, 4).Value = "峰段电量(P)";
                worksheet.Cell(1, 5).Value = "平段电量(T)";
                worksheet.Cell(1, 6).Value = "谷段电量(X)";
                worksheet.Cell(2, 1).Value = customer;
                worksheet.Cell(2, 2).Value = total;
                worksheet.Cell(2, 3).Value = sharp;
                worksheet.Cell(2, 4).Value = peak;
                worksheet.Cell(2, 5).Value = flat;
                worksheet.Cell(2, 6).Value = valley;

                Directory.CreateDirectory(Path.GetDirectoryName(path));
                workbook.SaveAs(path);
            }
        }

        private static void WriteRawDetailCsv(string path, params string[] dataLines)
        {
            var lines = new[] { "header1", "header2", "header3" }.Concat(dataLines).ToArray();
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            File.WriteAllLines(path, lines, Encoding.GetEncoding("GB18030"));
        }

        private static string RawDetailLine(
            string code,
            string name,
            double total,
            double sharp,
            double peak,
            double flat,
            double valley)
        {
            var columns = Enumerable.Repeat(string.Empty, 24).ToArray();
            columns[2] = code;
            columns[3] = name;
            columns[8] = total.ToString("0.####");
            columns[11] = sharp.ToString("0.####");
            columns[15] = peak.ToString("0.####");
            columns[19] = flat.ToString("0.####");
            columns[23] = valley.ToString("0.####");
            return string.Join(",", columns);
        }

        private static int FindCustomerRow(IXLWorksheet worksheet, string customer)
        {
            var row = TryFindCustomerRow(worksheet, customer);
            if (row > 0)
            {
                return row;
            }

            Assert.Fail("Customer not found: " + customer);
            return -1;
        }

        private static int TryFindCustomerRow(IXLWorksheet worksheet, string customer)
        {
            var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 1;
            for (var row = 4; row <= lastRow; row++)
            {
                if (worksheet.Cell(row, 3).GetString() == customer)
                {
                    return row;
                }
            }

            return -1;
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
