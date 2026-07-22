using System;
using System.IO;
using ClosedXML.Excel;
using HainanSettlementTool.Core.Models;
using HainanSettlementTool.Core.Services;

namespace HainanSettlementTool.Excel
{
    internal static class HainanStage2TemplateIndex
    {
        internal static HainanStage2TemplateCatalog Build(string proxyRoot, string interRoot)
        {
            var result = new HainanStage2TemplateCatalog();
            IndexTemplateRoot(result, "代理", proxyRoot);
            IndexTemplateRoot(result, "居间", interRoot);
            return result;
        }

        private static void IndexTemplateRoot(HainanStage2TemplateCatalog result, string kind, string root)
        {
            if (!Directory.Exists(root))
            {
                return;
            }

            foreach (var path in Directory.GetFiles(root, "*.xlsx", SearchOption.AllDirectories))
            {
                var fileName = Path.GetFileName(path);
                if (fileName.StartsWith("~$", StringComparison.OrdinalIgnoreCase)
                    || fileName.StartsWith("._", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                try
                {
                    using (var workbook = new XLWorkbook(path))
                    {
                        var sheet = HainanStage2ExcelUtil.LastMonthSheet(workbook);
                        var rawEntity = TextUtil.S(sheet.Cell("A2").GetFormattedString()).Replace("代理名称:", string.Empty);
                        if (string.IsNullOrWhiteSpace(rawEntity))
                        {
                            throw new InvalidDataException("模板最近月份工作表 A2 未填写主体名称。");
                        }

                        var owner = new DirectoryInfo(Path.GetDirectoryName(path)).Name.Replace(" - 海南2026", string.Empty);
                        result.Candidates.Add(new HainanStage2TemplateCandidate
                        {
                            Kind = kind,
                            Entity = rawEntity,
                            Owner = owner,
                            Path = path
                        });
                    }
                }
                catch (Exception ex)
                {
                    var settlementKind = kind + "费";
                    result.Issues.Add(new HainanStage2CheckIssue
                    {
                        Code = Stage2PreflightIssueKinds.TemplateUnreadable,
                        Disposition = Stage2PreflightDisposition.Blocker,
                        Severity = "错误",
                        Category = "上月分表模板无法读取",
                        Kind = settlementKind,
                        SettlementKind = settlementKind,
                        TemplateFile = path,
                        Message = "读取上月" + settlementKind + "分表模板失败：" + ex.Message,
                        Suggestion = "请修复或移出损坏、无关的工作簿后重新预检。"
                    });
                }
            }
        }
    }
}
