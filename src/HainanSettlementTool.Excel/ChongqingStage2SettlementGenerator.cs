using System.Collections.Generic;
using System.IO;
using HainanSettlementTool.Core.Models;

namespace HainanSettlementTool.Excel
{
    internal sealed class ChongqingStage2SettlementGenerator
    {
        public ChongqingStage2PreflightReport Analyze(ChongqingStage2Options options)
        {
            var details = ChongqingStage2LedgerReader.ReadDetails(options);
            var groups = ChongqingStage2LedgerReader.BuildGroups(details);
            var report = new ChongqingStage2PreflightReport { Month = options.Month };
            ChongqingStage2SummaryWorkbookWriter.AddSummaryPaymentIssues(options, groups, report.Issues);
            return report;
        }

        public ChongqingStage2Report Generate(ChongqingStage2Options options)
        {
            Directory.CreateDirectory(options.OutputDirectory);

            var details = ChongqingStage2LedgerReader.ReadDetails(options);
            var warnings = new List<string>();
            var auditIssues = new List<ChongqingStage2CheckIssue>();
            var groups = ChongqingStage2SplitWorkbookWriter.BuildSplitFiles(options, details, warnings, auditIssues);
            var summaryPath = ChongqingStage2SummaryWorkbookWriter.BuildSummary(options, groups, warnings);
            var report = ChongqingStage2ReportWriter.CreateReport(options, details, groups, summaryPath, warnings, auditIssues);
            ChongqingStage2ReportWriter.Write(report);
            return report;
        }
    }
}
