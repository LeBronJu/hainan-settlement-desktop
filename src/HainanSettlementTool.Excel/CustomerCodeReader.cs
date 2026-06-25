using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace HainanSettlementTool.Excel
{
    internal sealed class CustomerCodeReader
    {
        private readonly RawDetailRowReader _rowReader = new RawDetailRowReader();

        public Dictionary<string, string> Read(string rawDetailPath)
        {
            var result = new Dictionary<string, string>();
            if (string.IsNullOrWhiteSpace(rawDetailPath) || !File.Exists(rawDetailPath))
            {
                return result;
            }

            if (!RawDetailRowReader.IsSupported(rawDetailPath))
            {
                return result;
            }

            foreach (var row in _rowReader.Read(rawDetailPath, RawDetailSheetSelection.CustomerCodeSheets)
                .Where(row => row.Key.Length > 0 && row.CustomerCode.Length > 0))
            {
                if (!result.ContainsKey(row.Key))
                {
                    result[row.Key] = row.CustomerCode;
                }
            }

            return result;
        }
    }
}
