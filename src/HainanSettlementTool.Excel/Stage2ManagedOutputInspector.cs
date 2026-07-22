using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;

namespace HainanSettlementTool.Excel
{
    internal static class Stage2ManagedOutputInspector
    {
        internal static IReadOnlyList<Stage2ManagedWorkbookFinding> InspectUnexpectedWorkbooks(
            IEnumerable<string> managedRoots,
            IEnumerable<string> plannedWorkbookPaths,
            string targetSheetName)
        {
            if (string.IsNullOrWhiteSpace(targetSheetName))
            {
                throw new ArgumentException("请提供阶段二目标月份 sheet 名。", nameof(targetSheetName));
            }

            var findings = new List<Stage2ManagedWorkbookFinding>();
            var normalizedRoots = NormalizeRoots(managedRoots, findings);
            var planned = NormalizePlannedPaths(plannedWorkbookPaths, normalizedRoots, findings);

            foreach (var plannedPath in planned.Where(File.Exists))
            {
                InspectPlannedWorkbook(plannedPath, targetSheetName, findings);
            }

            foreach (var root in normalizedRoots)
            {
                if (Directory.Exists(root))
                {
                    InspectDirectory(root, planned, targetSheetName, findings);
                }
            }

            return findings
                .GroupBy(
                    finding => finding.Kind + "|" + finding.Path,
                    StringComparer.OrdinalIgnoreCase)
                .Select(group => group.First())
                .OrderBy(finding => finding.Path, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static List<string> NormalizeRoots(
            IEnumerable<string> managedRoots,
            IList<Stage2ManagedWorkbookFinding> findings)
        {
            var candidates = new List<string>();
            foreach (var rootValue in managedRoots ?? Enumerable.Empty<string>())
            {
                if (string.IsNullOrWhiteSpace(rootValue))
                {
                    continue;
                }

                try
                {
                    candidates.Add(TrimDirectorySeparator(Path.GetFullPath(rootValue)));
                }
                catch (Exception ex)
                {
                    findings.Add(Stage2ManagedWorkbookFinding.Unreadable(
                        rootValue,
                        "受管输出根路径无效，无法执行目标月残留检查：" + ex.Message));
                }
            }

            var distinct = candidates
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(path => path.Length)
                .ThenBy(path => path, StringComparer.OrdinalIgnoreCase)
                .ToList();
            return distinct
                .Where(candidate => !distinct.Any(other =>
                    !string.Equals(candidate, other, StringComparison.OrdinalIgnoreCase)
                    && IsStrictlyWithin(other, candidate)))
                .ToList();
        }

        private static HashSet<string> NormalizePlannedPaths(
            IEnumerable<string> plannedWorkbookPaths,
            IEnumerable<string> normalizedRoots,
            IList<Stage2ManagedWorkbookFinding> findings)
        {
            var roots = normalizedRoots.ToList();
            var planned = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var pathValue in plannedWorkbookPaths ?? Enumerable.Empty<string>())
            {
                if (string.IsNullOrWhiteSpace(pathValue))
                {
                    continue;
                }

                try
                {
                    var path = Path.GetFullPath(pathValue);
                    if (!roots.Any(root => IsStrictlyWithin(root, path)))
                    {
                        findings.Add(Stage2ManagedWorkbookFinding.Unreadable(
                            path,
                            "本批规划分表路径不在任何受管输出根内，已拒绝发布。"));
                        continue;
                    }

                    planned.Add(path);
                }
                catch (Exception ex)
                {
                    findings.Add(Stage2ManagedWorkbookFinding.Unreadable(
                        pathValue,
                        "本批规划分表路径无效：" + ex.Message));
                }
            }

            return planned;
        }

        private static void InspectDirectory(
            string directory,
            ISet<string> planned,
            string targetSheetName,
            IList<Stage2ManagedWorkbookFinding> findings)
        {
            try
            {
                if (IsReparsePoint(directory))
                {
                    findings.Add(Stage2ManagedWorkbookFinding.Unreadable(
                        directory,
                        "受管输出目录包含重解析目录，程序无法证明目标月文件集合完整。"));
                    return;
                }

                foreach (var path in Directory.GetFiles(directory, "*.xlsx", SearchOption.TopDirectoryOnly))
                {
                    InspectWorkbook(path, planned, targetSheetName, findings);
                }

                foreach (var child in Directory.GetDirectories(directory, "*", SearchOption.TopDirectoryOnly))
                {
                    InspectDirectory(child, planned, targetSheetName, findings);
                }
            }
            catch (Exception ex)
            {
                findings.Add(Stage2ManagedWorkbookFinding.Unreadable(
                    directory,
                    "受管输出目录无法完整扫描：" + ex.Message));
            }
        }

        private static void InspectWorkbook(
            string path,
            ISet<string> planned,
            string targetSheetName,
            IList<Stage2ManagedWorkbookFinding> findings)
        {
            var fullPath = Path.GetFullPath(path);
            if (planned.Contains(fullPath) || IsIgnoredWorkbook(path))
            {
                return;
            }

            try
            {
                if (IsReparsePoint(fullPath))
                {
                    findings.Add(Stage2ManagedWorkbookFinding.Unreadable(
                        fullPath,
                        "受管输出中存在重解析 workbook，程序无法安全检查。"));
                    return;
                }

                using (var stream = new FileStream(
                    fullPath,
                    FileMode.Open,
                    FileAccess.Read,
                    FileShare.ReadWrite | FileShare.Delete))
                using (var workbook = new XLWorkbook(stream))
                {
                    if (workbook.Worksheets.Any(sheet =>
                        string.Equals(sheet.Name, targetSheetName, StringComparison.OrdinalIgnoreCase)))
                    {
                        findings.Add(Stage2ManagedWorkbookFinding.UnexpectedTargetMonth(fullPath, targetSheetName));
                    }
                }
            }
            catch (Exception ex)
            {
                findings.Add(Stage2ManagedWorkbookFinding.Unreadable(
                    fullPath,
                    "非本批计划 workbook 无法读取，程序不能确认其中是否含目标月份 sheet：" + ex.Message));
            }
        }

        private static void InspectPlannedWorkbook(
            string path,
            string targetSheetName,
            IList<Stage2ManagedWorkbookFinding> findings)
        {
            try
            {
                if (IsReparsePoint(path))
                {
                    findings.Add(Stage2ManagedWorkbookFinding.Unreadable(
                        path,
                        "本批规划输出是重解析 workbook，已拒绝覆盖。"));
                    return;
                }

                using (var stream = new FileStream(
                    path,
                    FileMode.Open,
                    FileAccess.Read,
                    FileShare.ReadWrite | FileShare.Delete))
                using (var workbook = new XLWorkbook(stream))
                {
                    if (workbook.Worksheets.Any(sheet =>
                        string.Equals(sheet.Name, targetSheetName, StringComparison.OrdinalIgnoreCase)))
                    {
                        findings.Add(Stage2ManagedWorkbookFinding.PlannedTargetMonth(path, targetSheetName));
                    }
                }
            }
            catch (Exception ex)
            {
                findings.Add(Stage2ManagedWorkbookFinding.Unreadable(
                    path,
                    "本批规划输出已存在但无法读取，程序不能安全确认覆盖范围：" + ex.Message));
            }
        }

        private static bool IsIgnoredWorkbook(string path)
        {
            var name = Path.GetFileName(path) ?? string.Empty;
            return name.StartsWith("~$", StringComparison.OrdinalIgnoreCase)
                || name.StartsWith("._", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsReparsePoint(string path)
        {
            return (File.GetAttributes(path) & FileAttributes.ReparsePoint) != 0;
        }

        private static bool IsStrictlyWithin(string root, string candidate)
        {
            var prefix = TrimDirectorySeparator(Path.GetFullPath(root)) + Path.DirectorySeparatorChar;
            return Path.GetFullPath(candidate).StartsWith(prefix, StringComparison.OrdinalIgnoreCase);
        }

        private static string TrimDirectorySeparator(string path)
        {
            return path.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        }
    }

    internal enum Stage2ManagedWorkbookFindingKind
    {
        PlannedTargetMonth = 0,
        UnexpectedTargetMonth = 1,
        Unreadable = 2
    }

    internal sealed class Stage2ManagedWorkbookFinding
    {
        private Stage2ManagedWorkbookFinding(
            string path,
            Stage2ManagedWorkbookFindingKind kind,
            string message)
        {
            Path = path;
            Kind = kind;
            Message = message;
        }

        internal string Path { get; }

        internal Stage2ManagedWorkbookFindingKind Kind { get; }

        internal bool IsUnreadable
        {
            get { return Kind == Stage2ManagedWorkbookFindingKind.Unreadable; }
        }

        internal bool IsPlannedTargetMonth
        {
            get { return Kind == Stage2ManagedWorkbookFindingKind.PlannedTargetMonth; }
        }

        internal string Message { get; }

        internal static Stage2ManagedWorkbookFinding UnexpectedTargetMonth(
            string path,
            string targetSheetName)
        {
            return new Stage2ManagedWorkbookFinding(
                path,
                Stage2ManagedWorkbookFindingKind.UnexpectedTargetMonth,
                "非本批计划 workbook 仍含目标月份 sheet“" + targetSheetName
                    + "”，继续发布会把旧主体混入本批正式结果。请移走该文件、删除其中目标月 sheet，或改用新的输出目录。");
        }

        internal static Stage2ManagedWorkbookFinding PlannedTargetMonth(
            string path,
            string targetSheetName)
        {
            return new Stage2ManagedWorkbookFinding(
                path,
                Stage2ManagedWorkbookFindingKind.PlannedTargetMonth,
                "本批规划输出已存在目标月份 sheet“" + targetSheetName
                    + "”，程序将按同月重跑处理并由整包事务覆盖，请确认没有未回填到输入模板的人工修改。");
        }

        internal static Stage2ManagedWorkbookFinding Unreadable(string path, string message)
        {
            return new Stage2ManagedWorkbookFinding(
                path,
                Stage2ManagedWorkbookFindingKind.Unreadable,
                message);
        }
    }
}
