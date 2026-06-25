using System.Collections.Generic;
using System.Linq;
using HainanSettlementTool.Core.Models;

namespace HainanSettlementTool.Excel
{
    internal sealed class RawDetailReader
    {
        private readonly RawDetailRowReader _rowReader = new RawDetailRowReader();

        public List<PowerRow> Read(string rawDetailPath)
        {
            return _rowReader
                .Read(rawDetailPath, RawDetailSheetSelection.FirstSheet)
                .Where(row => row.HasPowerColumns)
                .Select(row => new PowerRow
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
