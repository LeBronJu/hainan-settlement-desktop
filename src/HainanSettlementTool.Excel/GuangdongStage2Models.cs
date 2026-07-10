using System;
using HainanSettlementTool.Core.Models;

namespace HainanSettlementTool.Excel
{
    internal sealed class GuangdongStage2WorkbookPreparation
    {
        public GuangdongStage2WorkbookPlan Plan { get; set; }
        public GuangdongStage2WorksheetLayout Layout { get; set; }
    }

    internal sealed class GuangdongStage2WorksheetLayout
    {
        public int HeaderRow { get; set; }
        public int DetailStartRow { get; set; }
        public int DetailEndRow { get; set; }
        public int TotalRow { get; set; }
        public string PeriodCellAddress { get; set; }
        public string SettlementDateCellAddress { get; set; }
        public string TargetPeriodText { get; set; }
        public string TargetSettlementDateText { get; set; }
    }

    internal sealed class GuangdongStage2DateField
    {
        public string CellAddress { get; set; }
        public string Text { get; set; }
        public DateTime Value { get; set; }
    }
}
