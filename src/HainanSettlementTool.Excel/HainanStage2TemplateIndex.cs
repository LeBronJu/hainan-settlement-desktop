using System;
using System.Collections.Generic;
using System.IO;
using ClosedXML.Excel;
using HainanSettlementTool.Core.Services;

namespace HainanSettlementTool.Excel
{
    internal static class HainanStage2TemplateIndex
    {
        internal static Dictionary<string, string> Build(string proxyRoot, string interRoot)
        {
            var result = new Dictionary<string, string>();
            IndexTemplateRoot(result, "代理", proxyRoot);
            IndexTemplateRoot(result, "居间", interRoot);
            return result;
        }

        private static void IndexTemplateRoot(IDictionary<string, string> result, string kind, string root)
        {
            if (!Directory.Exists(root))
            {
                return;
            }

            foreach (var path in Directory.GetFiles(root, "*.xlsx", SearchOption.AllDirectories))
            {
                if (Path.GetFileName(path).StartsWith("~$", StringComparison.Ordinal))
                {
                    continue;
                }

                try
                {
                    using (var workbook = new XLWorkbook(path))
                    {
                        var sheet = HainanStage2ExcelUtil.LastMonthSheet(workbook);
                        var rawEntity = TextUtil.S(sheet.Cell("A2").GetFormattedString()).Replace("代理名称:", string.Empty);
                        var owner = new DirectoryInfo(Path.GetDirectoryName(path)).Name.Replace(" - 海南2026", string.Empty);
                        result[HainanStage2ExcelUtil.TemplateKey(kind, owner, rawEntity)] = path;
                    }
                }
                catch
                {
                    // Ignore broken or unrelated workbooks in template folders.
                }
            }
        }
    }
}
