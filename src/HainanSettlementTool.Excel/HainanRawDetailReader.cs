using System.Collections.Generic;
using System.Linq;
using HainanSettlementTool.Core.Models;

namespace HainanSettlementTool.Excel
{
    internal sealed class HainanRawDetailReader
    {
        private readonly HainanRawDetailRowReader _rowReader = new HainanRawDetailRowReader();

        public List<HainanPowerRow> Read(string rawDetailPath)
        {
            return _rowReader
                .Read(rawDetailPath, HainanRawDetailSheetSelection.FirstSheet)
                .Where(row => row.HasPowerColumns)
                .Select(row => new HainanPowerRow
                {
                    SourceRow = row.SourceRow,
                    Name = row.Name,
                    Key = row.Key,
                    Total = row.Total,
                    Sharp = row.Sharp,
                    Peak = row.Peak,
                    Flat = row.Flat,
                    Valley = row.Valley
                })
                .ToList();
        }
    }
}
