using HainanSettlementTool.Core.Models;

namespace HainanSettlementTool.Excel
{
    internal sealed class GuangdongStage2MonthPreparationGenerator
    {
        private readonly GuangdongStage2WorkbookInspector _inspector = new GuangdongStage2WorkbookInspector();
        private readonly GuangdongStage2WorkbookWriter _writer = new GuangdongStage2WorkbookWriter();
        private readonly GuangdongStage2ReportWriter _reportWriter = new GuangdongStage2ReportWriter();

        public GuangdongStage2PreflightReport Analyze(GuangdongStage2MonthPreparationOptions options)
        {
            var preparations = _inspector.Analyze(options);
            var report = new GuangdongStage2PreflightReport
            {
                Year = options.Year,
                Month = options.Month
            };
            foreach (var preparation in preparations)
            {
                report.Workbooks.Add(preparation.Plan);
            }

            return report;
        }

        public GuangdongStage2MonthPreparationReport Generate(GuangdongStage2MonthPreparationOptions options)
        {
            var preparations = _inspector.Analyze(options);
            var runDirectory = GuangdongStage2ExcelUtil.CreateUniqueRunDirectory(
                options.OutputDirectory,
                options.Year,
                options.Month);
            var report = new GuangdongStage2MonthPreparationReport
            {
                Year = options.Year,
                Month = options.Month,
                OutputDirectory = runDirectory
            };
            report.Workbooks.AddRange(_writer.Write(runDirectory, preparations));
            _reportWriter.Write(report);
            return report;
        }
    }
}
