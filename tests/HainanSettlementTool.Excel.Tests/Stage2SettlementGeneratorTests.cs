using System;
using System.IO;
using ClosedXML.Excel;
using HainanSettlementTool.Core.Models;
using HainanSettlementTool.Core.Services;
using HainanSettlementTool.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace HainanSettlementTool.Excel.Tests
{
    [TestClass]
    public sealed class Stage2SettlementGeneratorTests
    {
        [TestMethod]
        public void GenerateSettlementExtendsDetailTotalsAndRepairsTemplateFormatting()
        {
            var root = Path.Combine(Path.GetTempPath(), "HainanSettlementToolTests", Guid.NewGuid().ToString("N"));
            var ledgerPath = Path.Combine(root, "ledger.xlsx");
            var proxyRoot = Path.Combine(root, "proxy");
            var interRoot = Path.Combine(root, "intermediary");
            var outputRoot = Path.Combine(root, "output");
            var summaryPath = Path.Combine(root, "summary.xlsx");

            try
            {
                Directory.CreateDirectory(proxyRoot);
                Directory.CreateDirectory(interRoot);
                Directory.CreateDirectory(outputRoot);

                WriteLedgerWithProxyRows(ledgerPath, "测试负责人", "测试代理", "存量客户", "新增客户");
                WriteProxyTemplate(proxyRoot, "测试负责人", "测试代理");
                WriteSummaryTemplate(summaryPath, "测试代理", "代理费", "清辉");

                var options = new Stage2Options
                {
                    Month = 4,
                    LedgerPath = ledgerPath,
                    ProxyTemplateDirectory = proxyRoot,
                    IntermediaryTemplateDirectory = interRoot,
                    SummaryTemplatePath = summaryPath,
                    OutputDirectory = outputRoot
                };

                new Stage2Service(new ClosedXmlStage1ExcelGateway()).Run(options, null);

                var outputPath = Path.Combine(
                    outputRoot,
                    "2026年代理 - 海南",
                    "测试负责人 - 海南2026",
                    "测试代理 2026海南.xlsx");
                using (var workbook = new XLWorkbook(outputPath))
                {
                    var worksheet = workbook.Worksheet("4月");
                    Assert.AreEqual("SUM(C5:C6)", worksheet.Cell(7, 3).FormulaA1);
                    Assert.AreEqual("SUM(P5:P6)", worksheet.Cell(7, 16).FormulaA1);
                    Assert.AreEqual("日期：2026年05月08日", FindSignatureDateText(worksheet));
                    AssertStyleMatches(workbook.Worksheet("2月").Cell(6, 3), worksheet.Cell(7, 3));
                    AssertStyleMatches(worksheet.Cell(5, 2), worksheet.Cell(6, 2));
                }
            }
            finally
            {
                if (Directory.Exists(root))
                {
                    Directory.Delete(root, true);
                }
            }
        }

        [TestMethod]
        public void GenerateSettlementShiftsBottomExcelDateCells()
        {
            var root = Path.Combine(Path.GetTempPath(), "HainanSettlementToolTests", Guid.NewGuid().ToString("N"));
            var ledgerPath = Path.Combine(root, "ledger.xlsx");
            var proxyRoot = Path.Combine(root, "proxy");
            var interRoot = Path.Combine(root, "intermediary");
            var outputRoot = Path.Combine(root, "output");
            var summaryPath = Path.Combine(root, "summary.xlsx");

            try
            {
                Directory.CreateDirectory(proxyRoot);
                Directory.CreateDirectory(interRoot);
                Directory.CreateDirectory(outputRoot);

                WriteLedgerWithProxyRows(ledgerPath, "测试负责人", "日期代理", "存量客户", "新增客户");
                WriteProxyTemplateWithExcelDateSignature(proxyRoot, "测试负责人", "日期代理");
                WriteSummaryTemplate(summaryPath, "日期代理", "代理费", "清辉");

                var options = new Stage2Options
                {
                    Month = 4,
                    LedgerPath = ledgerPath,
                    ProxyTemplateDirectory = proxyRoot,
                    IntermediaryTemplateDirectory = interRoot,
                    SummaryTemplatePath = summaryPath,
                    OutputDirectory = outputRoot
                };

                new Stage2Service(new ClosedXmlStage1ExcelGateway()).Run(options, null);

                var outputPath = Path.Combine(
                    outputRoot,
                    "2026年代理 - 海南",
                    "测试负责人 - 海南2026",
                    "日期代理 2026海南.xlsx");
                using (var workbook = new XLWorkbook(outputPath))
                {
                    var worksheet = workbook.Worksheet("4月");
                    var dateCell = FindFormattedCell(worksheet, "2026年6月8日");
                    Assert.AreEqual("yyyy\"年\"m\"月\"d\"日\";@", dateCell.Style.DateFormat.Format);
                }
            }
            finally
            {
                if (Directory.Exists(root))
                {
                    Directory.Delete(root, true);
                }
            }
        }

        [TestMethod]
        public void GenerateSettlementKeepsSummaryFooterOutOfDataRowsWhenAddingSubject()
        {
            var root = Path.Combine(Path.GetTempPath(), "HainanSettlementToolTests", Guid.NewGuid().ToString("N"));
            var ledgerPath = Path.Combine(root, "ledger.xlsx");
            var proxyRoot = Path.Combine(root, "proxy");
            var interRoot = Path.Combine(root, "intermediary");
            var outputRoot = Path.Combine(root, "output");
            var summaryPath = Path.Combine(root, "summary.xlsx");

            try
            {
                Directory.CreateDirectory(proxyRoot);
                Directory.CreateDirectory(interRoot);
                Directory.CreateDirectory(outputRoot);

                WriteLedgerWithProxyEntities(ledgerPath, "测试负责人", "存量代理", "新增代理");
                WriteProxyTemplate(proxyRoot, "测试负责人", "存量代理");
                WriteSummaryTemplateWithFooterInDataColumns(summaryPath, "存量代理", "代理费", "清辉");

                var options = new Stage2Options
                {
                    Month = 4,
                    LedgerPath = ledgerPath,
                    ProxyTemplateDirectory = proxyRoot,
                    IntermediaryTemplateDirectory = interRoot,
                    SummaryTemplatePath = summaryPath,
                    OutputDirectory = outputRoot
                };

                new Stage2Service(new ClosedXmlStage1ExcelGateway()).Run(options, null);

                var outputPath = Path.Combine(outputRoot, "【2026年海南省代理费汇总表-4月自动化】.xlsx");
                using (var workbook = new XLWorkbook(outputPath))
                {
                    var worksheet = workbook.Worksheet("汇总表");
                    Assert.AreEqual("存量代理", worksheet.Cell(4, 2).GetFormattedString());
                    Assert.AreEqual("新增代理", worksheet.Cell(5, 2).GetFormattedString());
                    Assert.AreEqual("合计", worksheet.Cell(6, 1).GetFormattedString());
                    Assert.AreEqual("当月实际支付", worksheet.Cell(2, 27).GetFormattedString());
                    Assert.AreEqual("SUM(V4:V5)", worksheet.Cell(6, 22).FormulaA1);
                    Assert.AreEqual("日期：2026年06月08日", worksheet.Cell(9, 2).GetFormattedString());
                    Assert.AreEqual(string.Empty, worksheet.Cell(9, 1).GetFormattedString());
                    AssertStyleMatches(worksheet.Cell(4, 2), worksheet.Cell(5, 2));
                    Assert.AreEqual("清辉", worksheet.Cell(5, 36).GetFormattedString());
                }
            }
            finally
            {
                if (Directory.Exists(root))
                {
                    Directory.Delete(root, true);
                }
            }
        }

        private static void WriteLedgerWithProxyRows(
            string path,
            string owner,
            string proxyEntity,
            string existingCustomer,
            string newCustomer)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.AddWorksheet(LedgerLayout.MainSheetName);
                var start = LedgerLayout.MonthStartColumn(4);
                worksheet.Cell(1, start).Value = "4月";
                WriteLedgerRow(worksheet, 4, start, owner, proxyEntity, existingCustomer, 100);
                WriteLedgerRow(worksheet, 5, start, owner, proxyEntity, newCustomer, 200);
                workbook.SaveAs(path);
            }
        }

        private static void WriteLedgerWithProxyEntities(
            string path,
            string owner,
            string existingEntity,
            string newEntity)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.AddWorksheet(LedgerLayout.MainSheetName);
                var start = LedgerLayout.MonthStartColumn(4);
                worksheet.Cell(1, start).Value = "4月";
                WriteLedgerRow(worksheet, 4, start, owner, existingEntity, "存量客户", 100);
                WriteLedgerRow(worksheet, 5, start, owner, newEntity, "新增主体客户", 200);
                workbook.SaveAs(path);
            }
        }

        private static void WriteLedgerRow(
            IXLWorksheet worksheet,
            int row,
            int start,
            string owner,
            string proxyEntity,
            string customer,
            double total)
        {
            worksheet.Cell(row, 3).Value = customer;
            worksheet.Cell(row, 8).Value = proxyEntity;
            worksheet.Cell(row, 10).Value = owner;
            worksheet.Cell(row, start).Value = total;
            worksheet.Cell(row, start + 1).Value = total * 0.1;
            worksheet.Cell(row, start + 2).Value = total * 0.2;
            worksheet.Cell(row, start + 3).Value = total * 0.3;
            worksheet.Cell(row, start + 4).Value = total * 0.4;
            worksheet.Cell(row, start + 5).Value = total * 0.5;
            worksheet.Cell(row, start + 6).Value = total * 0.6;
            worksheet.Cell(row, start + 13).Value = 0.5;
            worksheet.Cell(row, start + 14).Value = 1.2;
            worksheet.Cell(row, start + 16).Value = 0.06;
        }

        private static void WriteProxyTemplate(string root, string owner, string entity)
        {
            var folder = Path.Combine(root, owner + " - 海南2026");
            Directory.CreateDirectory(folder);
            var path = Path.Combine(folder, entity + " 2026海南.xlsx");
            using (var workbook = new XLWorkbook())
            {
                WriteDetailTemplateSheet(workbook.AddWorksheet("2月"), "代理", entity, 2, true);
                WriteDetailTemplateSheet(workbook.AddWorksheet("3月"), "代理", entity, 3, false);
                workbook.SaveAs(path);
            }
        }

        private static void WriteProxyTemplateWithExcelDateSignature(string root, string owner, string entity)
        {
            var folder = Path.Combine(root, owner + " - 海南2026");
            Directory.CreateDirectory(folder);
            var path = Path.Combine(folder, entity + " 2026海南.xlsx");
            using (var workbook = new XLWorkbook())
            {
                WriteDetailTemplateSheet(workbook.AddWorksheet("2月"), "代理", entity, 2, true);
                var sheet = workbook.AddWorksheet("3月");
                WriteDetailTemplateSheet(sheet, "代理", entity, 3, true);
                sheet.Cell("B8").Clear(XLClearOptions.Contents);
                sheet.Cell("N9").Value = new DateTime(2026, 5, 8);
                sheet.Cell("N9").Style.DateFormat.Format = "yyyy\"年\"m\"月\"d\"日\";@";
                sheet.Range("N9:O9").Merge();
                workbook.SaveAs(path);
            }
        }

        private static void WriteDetailTemplateSheet(
            IXLWorksheet worksheet,
            string kind,
            string entity,
            int month,
            bool styleTotalFormulaCells)
        {
            worksheet.Cell("A1").Value = kind + "费用结算单";
            worksheet.Cell("A2").Value = "代理名称:" + entity;
            worksheet.Cell("F2").Value = "所属期：2026年" + month.ToString("00") + "月";
            worksheet.Cell("M2").Value = "结算日期：2026 年 " + (month + 1).ToString("00") + " 月 15 日";

            worksheet.Cell(5, 1).Value = 1;
            worksheet.Cell(5, 2).Value = "存量客户";
            ApplyDetailRowStyle(worksheet.Row(5));

            worksheet.Cell(6, 1).Value = "合计";
            for (var column = 3; column <= 7; column++)
            {
                var letter = ColumnLetter(column);
                worksheet.Cell(6, column).FormulaA1 = "SUM(" + letter + "5:" + letter + "5)";
            }

            for (var column = 12; column <= 16; column++)
            {
                var letter = ColumnLetter(column);
                worksheet.Cell(6, column).FormulaA1 = "SUM(" + letter + "5:" + letter + "5)";
            }

            worksheet.Cell(8, 2).Value = "日期：2026年" + (month + 1).ToString("00") + "月08日";

            if (styleTotalFormulaCells)
            {
                ApplyTotalFormulaStyle(worksheet.Row(6));
            }
        }

        private static void ApplyDetailRowStyle(IXLRow row)
        {
            row.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            row.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            row.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            row.Style.NumberFormat.Format = "0.0000";
        }

        private static void ApplyTotalFormulaStyle(IXLRow row)
        {
            row.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            row.Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
            row.Style.Border.InsideBorder = XLBorderStyleValues.Medium;
            row.Style.NumberFormat.Format = "0.00";
        }

        private static void WriteSummaryTemplate(string path, string entity, string kind, string paymentParty)
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.AddWorksheet("汇总表");
                WriteSummaryHeaders(worksheet);
                worksheet.Cell(2, 22).Value = "累计代理费总计";
                worksheet.Cell(4, 1).Value = 1;
                worksheet.Cell(4, 2).Value = entity;
                worksheet.Cell(4, 3).Value = kind;
                worksheet.Cell(4, 30).Value = paymentParty;
                worksheet.Cell(5, 1).Value = "合计";
                worksheet.Cell(8, 31).Value = "日期：2026年05月08日";
                workbook.SaveAs(path);
            }
        }

        private static void WriteSummaryTemplateWithFooterInDataColumns(string path, string entity, string kind, string paymentParty)
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.AddWorksheet("汇总表");
                WriteSummaryHeaders(worksheet);
                worksheet.Cell(2, 22).Value = "累计代理费总计";
                worksheet.Cell(4, 1).Value = 1;
                worksheet.Cell(4, 2).Value = entity;
                worksheet.Cell(4, 3).Value = kind;
                worksheet.Cell(4, 30).Value = paymentParty;
                ApplySummaryDataRowStyle(worksheet.Row(4));
                worksheet.Cell(5, 1).Value = "合计";
                worksheet.Cell(8, 2).Value = "日期：2026年05月08日";
                workbook.SaveAs(path);
            }
        }

        private static void WriteSummaryHeaders(IXLWorksheet worksheet)
        {
            worksheet.Range(2, 16, 2, 20).Merge();
            worksheet.Cell(2, 16).Value = "2026年3月";
            worksheet.Cell(3, 16).Value = "代理费";
            worksheet.Cell(3, 17).Value = "居间费";
            worksheet.Cell(3, 18).Value = "退补电费";
            worksheet.Cell(3, 19).Value = "当月抵扣";
            worksheet.Cell(3, 20).Value = "费用合计";
            worksheet.Range(2, 21, 3, 21).Merge();
            worksheet.Cell(2, 21).Value = "当月实际支付";
        }

        private static void ApplySummaryDataRowStyle(IXLRow row)
        {
            row.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            row.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            row.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            row.Style.NumberFormat.Format = "0.00";
        }

        private static string FindSignatureDateText(IXLWorksheet worksheet)
        {
            foreach (var cell in worksheet.CellsUsed())
            {
                var value = cell.GetFormattedString();
                if (value.Contains("日期：") && !value.Contains("结算日期"))
                {
                    return value;
                }
            }

            Assert.Fail("未找到签字日期单元格。");
            return null;
        }

        private static IXLCell FindFormattedCell(IXLWorksheet worksheet, string formattedText)
        {
            foreach (var cell in worksheet.CellsUsed())
            {
                if (cell.GetFormattedString() == formattedText)
                {
                    return cell;
                }
            }

            Assert.Fail("未找到显示文本为“" + formattedText + "”的单元格。");
            return null;
        }

        private static void AssertStyleMatches(IXLCell expected, IXLCell actual)
        {
            Assert.AreEqual(expected.Style.Alignment.Horizontal, actual.Style.Alignment.Horizontal);
            Assert.AreEqual(expected.Style.Border.LeftBorder, actual.Style.Border.LeftBorder);
            Assert.AreEqual(expected.Style.Border.RightBorder, actual.Style.Border.RightBorder);
            Assert.AreEqual(expected.Style.Border.TopBorder, actual.Style.Border.TopBorder);
            Assert.AreEqual(expected.Style.Border.BottomBorder, actual.Style.Border.BottomBorder);
            Assert.AreEqual(expected.Style.NumberFormat.Format, actual.Style.NumberFormat.Format);
        }

        private static string ColumnLetter(int columnNumber)
        {
            var dividend = columnNumber;
            var columnName = string.Empty;
            while (dividend > 0)
            {
                var modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }

            return columnName;
        }
    }
}
