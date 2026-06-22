using System;
using System.IO;

namespace HainanSettlementTool.Core.Services
{
    public static class FileAccessGuard
    {
        public static void RequireReadableWorkbook(string path, string label)
        {
            RequireExistingFile(path, label);
            TryOpen(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite, label, "读取");
        }

        public static void RequireWritableWorkbook(string path, string label)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return;
            }

            if (!File.Exists(path))
            {
                var directory = Path.GetDirectoryName(path);
                if (!string.IsNullOrWhiteSpace(directory))
                {
                    Directory.CreateDirectory(directory);
                }
                return;
            }

            TryOpen(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None, label, "写入");
        }

        public static void RequireExistingFile(string path, string label)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                throw new ArgumentException("请选择" + label + "。");
            }

            if (Path.GetFileName(path).StartsWith("~$", StringComparison.Ordinal))
            {
                throw new ArgumentException(label + "选到了 Excel 临时文件，请选择正式文件。");
            }

            if (!File.Exists(path))
            {
                throw new FileNotFoundException(label + "不存在。", path);
            }
        }

        private static void TryOpen(string path, FileMode mode, FileAccess access, FileShare share, string label, string action)
        {
            try
            {
                using (File.Open(path, mode, access, share))
                {
                }
            }
            catch (IOException ex)
            {
                throw new IOException(label + "当前可能正被 Excel 或其他程序占用，无法" + action + "。请关闭该文件后重试。\n\n文件：" + path, ex);
            }
            catch (UnauthorizedAccessException ex)
            {
                throw new UnauthorizedAccessException(label + "没有" + action + "权限。请检查文件权限或换一个输出目录。\n\n文件：" + path, ex);
            }
        }
    }
}
