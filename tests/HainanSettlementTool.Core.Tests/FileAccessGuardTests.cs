using System;
using System.IO;
using HainanSettlementTool.Core.Services;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace HainanSettlementTool.Core.Tests
{
    [TestClass]
    public sealed class FileAccessGuardTests
    {
        [TestMethod]
        public void RequireWritableWorkbookCreatesParentDirectoryForNewOutput()
        {
            var root = Path.Combine(Path.GetTempPath(), "HainanSettlementToolTests", Guid.NewGuid().ToString("N"));
            var outputPath = Path.Combine(root, "nested", "output.xlsx");

            try
            {
                FileAccessGuard.RequireWritableWorkbook(outputPath, "输出文件");

                Assert.IsTrue(Directory.Exists(Path.GetDirectoryName(outputPath)));
                Assert.IsFalse(File.Exists(outputPath));
            }
            finally
            {
                if (Directory.Exists(root))
                {
                    Directory.Delete(root, true);
                }
            }
        }

        [TestMethod]
        public void RequireExistingFileRejectsExcelTemporaryFiles()
        {
            var path = Path.Combine(Path.GetTempPath(), "~$台账.xlsx");

            var ex = Assert.ThrowsException<ArgumentException>(() =>
                FileAccessGuard.RequireExistingFile(path, "基础台账"));

            StringAssert.Contains(ex.Message, "Excel 临时文件");
        }

        [TestMethod]
        public void RequireWritableWorkbookReportsLockedExistingOutput()
        {
            var root = Path.Combine(Path.GetTempPath(), "HainanSettlementToolTests", Guid.NewGuid().ToString("N"));
            var outputPath = Path.Combine(root, "output.xlsx");

            try
            {
                Directory.CreateDirectory(root);
                File.WriteAllText(outputPath, "locked");

                using (File.Open(outputPath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                {
                    var ex = Assert.ThrowsException<IOException>(() =>
                        FileAccessGuard.RequireWritableWorkbook(outputPath, "输出文件"));

                    StringAssert.Contains(ex.Message, "被 Excel 或其他程序占用");
                }
            }
            finally
            {
                if (Directory.Exists(root))
                {
                    Directory.Delete(root, true);
                }
            }
        }
    }
}
