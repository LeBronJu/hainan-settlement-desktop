using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using HainanSettlementTool.Core.Services;

namespace HainanSettlementTool.Excel
{
    internal static class Stage2BatchIntegrityVerifier
    {
        internal static void VerifyFiles(
            Stage2BatchWorkspace workspace,
            IEnumerable<string> splitWorkbookPaths,
            string summaryWorkbookPath,
            IEnumerable<string> reportPaths)
        {
            if (workspace == null)
            {
                throw new ArgumentNullException(nameof(workspace));
            }

            var splitPaths = (splitWorkbookPaths ?? Enumerable.Empty<string>()).ToList();
            var distinctSplitPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var path in splitPaths)
            {
                RequireStagedWorkbook(workspace, path, "阶段二分表");
                if (!distinctSplitPaths.Add(Path.GetFullPath(path)))
                {
                    throw new InvalidDataException("同一阶段二批次有多个汇总主体指向同一分表：" + path);
                }
            }

            RequireStagedWorkbook(workspace, summaryWorkbookPath, "阶段二汇总表");
            if (distinctSplitPaths.Contains(Path.GetFullPath(summaryWorkbookPath)))
            {
                throw new InvalidDataException("阶段二汇总表不能与分表使用同一输出路径。");
            }

            foreach (var path in reportPaths ?? Enumerable.Empty<string>())
            {
                RequireStagedFile(workspace, path, "阶段二报告");
            }
        }

        private static void RequireStagedWorkbook(
            Stage2BatchWorkspace workspace,
            string path,
            string label)
        {
            RequireStagedFile(workspace, path, label);
            try
            {
                using (var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.Read))
                using (var workbook = new XLWorkbook(stream))
                {
                    if (workbook.Worksheets.Count == 0)
                    {
                        throw new InvalidDataException(label + "没有工作表：" + path);
                    }
                }
            }
            catch (InvalidDataException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new InvalidDataException(label + "无法重新打开：" + path, ex);
            }
        }

        private static void RequireStagedFile(
            Stage2BatchWorkspace workspace,
            string path,
            string label)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                throw new InvalidDataException(label + "路径为空。");
            }

            workspace.GetFinalPath(path);
            if (!File.Exists(path))
            {
                throw new FileNotFoundException(label + "未生成。", path);
            }

            if (new FileInfo(path).Length <= 0)
            {
                throw new InvalidDataException(label + "是空文件：" + path);
            }
        }
    }
}
