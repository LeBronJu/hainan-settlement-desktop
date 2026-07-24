using System;
using System.IO;
using System.Linq;
using System.Text;
using ClosedXML.Excel;
using HainanSettlementTool.Core.Models;
using HainanSettlementTool.Core.Services;
using HainanSettlementTool.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace HainanSettlementTool.Excel.Tests
{
    [TestClass]
    public sealed class HainanStage1LedgerUpdaterTests
    {
        [TestMethod]
        public void HainanRunInsertsNewCustomerBeforeFooterAndWritesMonthPower()
        {
            var root = CreateTempRoot();
            var baseLedgerPath = Path.Combine(root, "base-ledger.xlsx");
            var powerPath = Path.Combine(root, "power.xlsx");
            var outputDirectory = Path.Combine(root, "out");

            try
            {
                WriteHainanBaseLedger(baseLedgerPath);
                WriteHainanPowerWorkbook(powerPath, "新增客户", 12.3456, 1.1, 2.2, 3.3, 4.4);

                var report = RunHainanStage1(baseLedgerPath, powerPath, outputDirectory);

                Assert.AreEqual(1, report.NewRows);
                Assert.AreEqual(1, report.NewCustomers.Count);
                Assert.AreEqual("新增客户", report.NewCustomers[0].Name);
                Assert.AreEqual(5, report.NewCustomers[0].TargetRow);

                using (var workbook = new XLWorkbook(report.Output))
                {
                    var worksheet = workbook.Worksheet(HainanLedgerLayout.MainSheetName);
                    var targetStart = HainanLedgerLayout.MonthStartColumn(4);

                    Assert.AreEqual("既有客户", worksheet.Cell(4, 3).GetString());
                    Assert.AreEqual("新增客户", worksheet.Cell(5, 3).GetString());
                    Assert.AreEqual("合计", worksheet.Cell(6, 1).GetString());
                    Assert.AreEqual(12.3456, worksheet.Cell(5, targetStart).GetDouble(), 0.00001);
                    Assert.AreEqual(1.1, worksheet.Cell(5, targetStart + 1).GetDouble(), 0.00001);
                    Assert.AreEqual(2.2, worksheet.Cell(5, targetStart + 2).GetDouble(), 0.00001);
                    Assert.AreEqual(3.3, worksheet.Cell(5, targetStart + 3).GetDouble(), 0.00001);
                    Assert.AreEqual(4.4, worksheet.Cell(5, targetStart + 4).GetDouble(), 0.00001);
                    Assert.IsTrue(worksheet.Cell(4, targetStart).IsEmpty());
                    Assert.IsTrue(worksheet.Cell(4, targetStart + 1).IsEmpty());
                    Assert.IsTrue(worksheet.Cell(4, targetStart + 2).IsEmpty());
                    Assert.IsTrue(worksheet.Cell(4, targetStart + 3).IsEmpty());
                    Assert.IsTrue(worksheet.Cell(4, targetStart + 4).IsEmpty());
                    Assert.IsTrue(worksheet.Cell(5, 23).IsEmpty());
                    Assert.IsTrue(worksheet.Cell(5, 24).IsEmpty());
                    Assert.IsTrue(worksheet.Cell(5, 25).IsEmpty());
                    Assert.IsTrue(worksheet.Cell(5, 27).IsEmpty());
                    Assert.IsTrue(worksheet.Cell(5, targetStart + 5).IsEmpty());
                    Assert.IsTrue(worksheet.Cell(5, targetStart + 6).HasFormula);
                }
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void HainanRunFillsOnlyUnambiguousCustomerCodesFromRawDetail()
        {
            var root = CreateTempRoot();
            var baseLedgerPath = Path.Combine(root, "base-ledger.xlsx");
            var rawDetailPath = Path.Combine(root, "raw.csv");
            var cleanPowerPath = Path.Combine(root, "out", "clean-power.xlsx");
            var outputDirectory = Path.Combine(root, "out");

            try
            {
                WriteHainanBaseLedger(baseLedgerPath);
                WriteHainanRawDetailCsv(
                    rawDetailPath,
                    HainanRawDetailLine("CODE-ONLY", "唯一客户", 10, 1, 2, 3, 4),
                    HainanRawDetailLine("CODE-A", "冲突客户", 5, 1, 1, 1, 2),
                    HainanRawDetailLine("CODE-B", "冲突客户", 6, 2, 2, 1, 1));

                var report = RunHainanStage1(
                    new HainanStage1Options
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
                    var worksheet = workbook.Worksheet(HainanLedgerLayout.MainSheetName);
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
        public void HainanRunWritesLedgerCopyWithoutChangingBaseLedger()
        {
            var root = CreateTempRoot();
            var baseLedgerPath = Path.Combine(root, "base-ledger.xlsx");
            var powerPath = Path.Combine(root, "power.xlsx");
            var outputDirectory = Path.Combine(root, "out");

            try
            {
                WriteHainanBaseLedger(baseLedgerPath);
                WriteHainanPowerWorkbook(powerPath, "新增客户", 12.3456, 1.1, 2.2, 3.3, 4.4);

                var report = RunHainanStage1(baseLedgerPath, powerPath, outputDirectory);

                Assert.AreNotEqual(Path.GetFullPath(baseLedgerPath), Path.GetFullPath(report.Output));
                using (var baseWorkbook = new XLWorkbook(baseLedgerPath))
                using (var outputWorkbook = new XLWorkbook(report.Output))
                {
                    var baseSheet = baseWorkbook.Worksheet(HainanLedgerLayout.MainSheetName);
                    var outputSheet = outputWorkbook.Worksheet(HainanLedgerLayout.MainSheetName);

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

        [TestMethod]
        public void HainanRunClearsAllExistingTargetMonthPowerBeforeWriting()
        {
            var root = CreateTempRoot();
            var baseLedgerPath = Path.Combine(root, "base-ledger.xlsx");
            var powerPath = Path.Combine(root, "power.xlsx");
            var outputDirectory = Path.Combine(root, "out");

            try
            {
                WriteHainanBaseLedger(baseLedgerPath);
                SeedHainanTargetMonthPower(baseLedgerPath, 4, 77, 7, 17, 23, 30);
                WriteHainanPowerWorkbook(powerPath, "新增客户", 12.3456, 1.1, 2.2, 3.3, 4.4);

                var report = RunHainanStage1(baseLedgerPath, powerPath, outputDirectory);

                Assert.IsTrue(report.TargetMonthAlreadyPresent);
                using (var workbook = new XLWorkbook(report.Output))
                {
                    var worksheet = workbook.Worksheet(HainanLedgerLayout.MainSheetName);
                    var targetStart = HainanLedgerLayout.MonthStartColumn(4);
                    for (var offset = 0; offset < 5; offset++)
                    {
                        Assert.IsTrue(
                            worksheet.Cell(4, targetStart + offset).IsEmpty(),
                            "台账独有客户的既有目标月电量应在写入前清空，偏移：" + offset);
                    }
                }
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }


        private static HainanStage1Report RunHainanStage1(string baseLedgerPath, string powerPath, string outputDirectory)
        {
            return RunHainanStage1(
                new HainanStage1Options
                {
                    Month = 4,
                    BaseLedgerPath = baseLedgerPath,
                    PowerPath = powerPath,
                    OutputDirectory = outputDirectory,
                    OutputLedgerName = "updated-ledger.xlsx"
                });
        }

        private static HainanStage1Report RunHainanStage1(HainanStage1Options options)
        {
            var service = new HainanStage1Service(new ClosedXmlSettlementExcelGateway());
            return service.Run(options, null);
        }

        private static void WriteHainanBaseLedger(string path)
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.AddWorksheet(HainanLedgerLayout.MainSheetName);
                var marchStart = HainanLedgerLayout.MonthStartColumn(3);

                worksheet.Cell(1, 1).Value = "售电结算台账";
                worksheet.Cell(1, marchStart).Value = "3月";
                worksheet.Cell(4, 1).Value = 1;
                worksheet.Cell(4, 2).Value = "OLD-CODE";
                worksheet.Cell(4, 3).Value = "既有客户";
                worksheet.Cell(4, 10).Value = "负责人";
                worksheet.Cell(4, 23).Value = "不得继承的合同值";
                worksheet.Cell(4, 24).Value = "不得继承的税务值";
                worksheet.Cell(4, 25).Value = "不得继承的代理值";
                worksheet.Cell(4, 27).Value = "不得继承的人工值";
                worksheet.Cell(4, marchStart).Value = 66;
                worksheet.Cell(4, marchStart + 1).Value = 6;
                worksheet.Cell(4, marchStart + 2).Value = 16;
                worksheet.Cell(4, marchStart + 3).Value = 20;
                worksheet.Cell(4, marchStart + 4).Value = 24;
                worksheet.Cell(4, marchStart + 5).Value = 1.2;
                worksheet.Cell(4, marchStart + 6).FormulaA1 = "=A4";
                worksheet.Cell(5, 1).Value = "合计";
                worksheet.Cell(5, 3).Value = "footer";

                Directory.CreateDirectory(Path.GetDirectoryName(path));
                workbook.SaveAs(path);
            }
        }

        private static void SeedHainanTargetMonthPower(
            string path,
            int month,
            double total,
            double sharp,
            double peak,
            double flat,
            double valley)
        {
            using (var workbook = new XLWorkbook(path))
            {
                var worksheet = workbook.Worksheet(HainanLedgerLayout.MainSheetName);
                var start = HainanLedgerLayout.MonthStartColumn(month);
                worksheet.Cell(1, start).Value = month + "月";
                worksheet.Cell(4, start).Value = total;
                worksheet.Cell(4, start + 1).Value = sharp;
                worksheet.Cell(4, start + 2).Value = peak;
                worksheet.Cell(4, start + 3).Value = flat;
                worksheet.Cell(4, start + 4).Value = valley;
                worksheet.Cell(4, start + 5).Value = 1.1;
                worksheet.Cell(4, start + 6).FormulaA1 = "=A4";
                workbook.Save();
            }
        }

        private static void WriteHainanPowerWorkbook(
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

        private static void WriteHainanRawDetailCsv(string path, params string[] dataLines)
        {
            var lines = new[] { "header1", "header2", "header3" }.Concat(dataLines).ToArray();
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            File.WriteAllLines(path, lines, Encoding.GetEncoding("GB18030"));
        }

        private static string HainanRawDetailLine(
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
