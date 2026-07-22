using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace HainanSettlementTool.Excel
{
    internal static class Stage2InputFingerprint
    {
        internal static string Capture(
            IEnumerable<string> explicitWorkbookPaths,
            IEnumerable<string> templateDirectories,
            IEnumerable<string> configurationTokens,
            IEnumerable<string> observedOutputPaths = null)
        {
            var paths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var path in explicitWorkbookPaths ?? Enumerable.Empty<string>())
            {
                if (!string.IsNullOrWhiteSpace(path))
                {
                    paths.Add(Path.GetFullPath(path));
                }
            }

            var directories = (templateDirectories ?? Enumerable.Empty<string>())
                .Where(directory => !string.IsNullOrWhiteSpace(directory))
                .Select(Path.GetFullPath)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(directory => directory, StringComparer.OrdinalIgnoreCase)
                .ToList();
            foreach (var directory in directories)
            {
                CollectTemplateWorkbooks(directory, paths);
            }

            var tokens = paths
                .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
                .Select(FileToken)
                .ToList();
            tokens.AddRange(directories.Select(directory => "directory|" + Encode(directory)));
            tokens.AddRange((configurationTokens ?? Enumerable.Empty<string>())
                .Select(value => "configuration|" + Encode(value))
                .OrderBy(value => value, StringComparer.Ordinal));
            tokens.AddRange((observedOutputPaths ?? Enumerable.Empty<string>())
                .Where(path => !string.IsNullOrWhiteSpace(path))
                .Select(Path.GetFullPath)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
                .Select(ObservedOutputToken));
            using (var sha256 = SHA256.Create())
            {
                var payload = string.Join("\n", tokens);
                var bytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(payload));
                return string.Concat(bytes.Select(value =>
                    value.ToString("x2", CultureInfo.InvariantCulture)));
            }
        }

        internal static bool Matches(string expected, string actual)
        {
            if (string.IsNullOrWhiteSpace(expected) || string.IsNullOrWhiteSpace(actual))
            {
                return false;
            }

            var left = Encoding.ASCII.GetBytes(expected.Trim());
            var right = Encoding.ASCII.GetBytes(actual.Trim());
            if (left.Length != right.Length)
            {
                return false;
            }

            var difference = 0;
            for (var index = 0; index < left.Length; index++)
            {
                difference |= left[index] ^ right[index];
            }

            return difference == 0;
        }

        private static void CollectTemplateWorkbooks(string directory, ISet<string> paths)
        {
            if (!Directory.Exists(directory))
            {
                throw new DirectoryNotFoundException("阶段二模板目录不存在：" + directory);
            }

            if (IsReparsePoint(directory))
            {
                throw new IOException("阶段二模板目录不能是重解析目录：" + directory);
            }

            foreach (var path in Directory.GetFiles(directory, "*.xlsx", SearchOption.TopDirectoryOnly))
            {
                if (!IsIgnoredWorkbook(path))
                {
                    paths.Add(Path.GetFullPath(path));
                }
            }

            foreach (var child in Directory.GetDirectories(directory, "*", SearchOption.TopDirectoryOnly))
            {
                CollectTemplateWorkbooks(child, paths);
            }
        }

        private static string FileToken(string path)
        {
            if (!File.Exists(path))
            {
                throw new FileNotFoundException("阶段二预检输入在生成前已不存在。", path);
            }

            if (IsReparsePoint(path))
            {
                throw new IOException("阶段二预检输入不能是重解析文件：" + path);
            }

            byte[] contentHash;
            using (var stream = new FileStream(
                path,
                FileMode.Open,
                FileAccess.Read,
                FileShare.ReadWrite | FileShare.Delete))
            using (var sha256 = SHA256.Create())
            {
                contentHash = sha256.ComputeHash(stream);
            }

            var info = new FileInfo(path);
            info.Refresh();
            return Encode(path)
                + Encode(info.Length.ToString(CultureInfo.InvariantCulture))
                + Encode(info.LastWriteTimeUtc.Ticks.ToString(CultureInfo.InvariantCulture))
                + Encode(string.Concat(contentHash.Select(value =>
                    value.ToString("x2", CultureInfo.InvariantCulture))));
        }

        private static string ObservedOutputToken(string path)
        {
            if (Directory.Exists(path))
            {
                if (IsReparsePoint(path))
                {
                    throw new IOException("阶段二规划输出不能是重解析目录：" + path);
                }

                return "observed-output|directory|" + Encode(path);
            }

            if (!File.Exists(path))
            {
                return "observed-output|missing|" + Encode(path);
            }

            return "observed-output|file|" + FileToken(path);
        }

        private static string Encode(string value)
        {
            var text = value ?? string.Empty;
            return text.Length.ToString(CultureInfo.InvariantCulture) + ":" + text + "|";
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
    }
}
