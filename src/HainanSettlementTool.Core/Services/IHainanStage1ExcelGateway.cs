using System.Collections.Generic;
using HainanSettlementTool.Core.Models;

namespace HainanSettlementTool.Core.Services
{
    public interface IHainanStage1ExcelGateway
    {
        List<HainanPowerRow> ReadPowerRows(string powerPath);
        List<HainanPowerRow> ReadRawPowerRows(string rawDetailPath);
        Dictionary<string, string> ReadCustomerCodes(string rawDetailPath);
        void WritePowerWorkbook(IEnumerable<HainanPowerRow> rows, string outputPath);
        HainanStage1Report UpdateLedger(HainanStage1Options options);
    }
}
