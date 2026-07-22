using System;
using System.IO;
using HainanSettlementTool.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace HainanSettlementTool.Excel.Tests
{
    [TestClass]
    public sealed class Stage2InputFingerprintTests
    {
        [TestMethod]
        public void CaptureChangesWhenExplicitInputChanges()
        {
            WithTempRoot(root =>
            {
                var input = Write(root, "ledger.xlsx", "before");
                var templates = Directory.CreateDirectory(Path.Combine(root, "templates")).FullName;
                var before = Capture(input, templates, "month|5", null);

                File.WriteAllText(input, "after");
                var after = Capture(input, templates, "month|5", null);

                Assert.AreNotEqual(before, after);
            });
        }

        [TestMethod]
        public void CaptureChangesWhenNestedTemplateOrConfigurationChanges()
        {
            WithTempRoot(root =>
            {
                var input = Write(root, "ledger.xlsx", "ledger");
                var templates = Directory.CreateDirectory(Path.Combine(root, "templates", "owner")).Parent.FullName;
                var template = Write(Path.Combine(templates, "owner"), "template.xlsx", "before");
                var before = Capture(input, templates, "month|5", null);

                File.WriteAllText(template, "after");
                var templateChanged = Capture(input, templates, "month|5", null);
                var configurationChanged = Capture(input, templates, "month|6", null);

                Assert.AreNotEqual(before, templateChanged);
                Assert.AreNotEqual(templateChanged, configurationChanged);
            });
        }

        [TestMethod]
        public void CaptureBindsMissingAndExistingPlannedOutputState()
        {
            WithTempRoot(root =>
            {
                var input = Write(root, "ledger.xlsx", "ledger");
                var templates = Directory.CreateDirectory(Path.Combine(root, "templates")).FullName;
                var planned = Path.Combine(root, "out", "owner.xlsx");
                var missing = Capture(input, templates, "month|5", planned);

                Write(Path.GetDirectoryName(planned), Path.GetFileName(planned), "generated then edited");
                var existing = Capture(input, templates, "month|5", planned);
                File.WriteAllText(planned, "second edit");
                var edited = Capture(input, templates, "month|5", planned);

                Assert.AreNotEqual(missing, existing);
                Assert.AreNotEqual(existing, edited);
            });
        }

        [TestMethod]
        public void CaptureIgnoresExcelTemporaryFilesInTemplateTree()
        {
            WithTempRoot(root =>
            {
                var input = Write(root, "ledger.xlsx", "ledger");
                var templates = Directory.CreateDirectory(Path.Combine(root, "templates")).FullName;
                var before = Capture(input, templates, "month|5", null);

                Write(templates, "~$editing.xlsx", "temporary");
                Write(templates, "._metadata.xlsx", "temporary");
                var after = Capture(input, templates, "month|5", null);

                Assert.AreEqual(before, after);
            });
        }

        private static string Capture(string input, string templates, string configuration, string planned)
        {
            return Stage2InputFingerprint.Capture(
                new[] { input },
                new[] { templates },
                new[] { configuration },
                planned == null ? null : new[] { planned });
        }

        private static string Write(string directory, string name, string content)
        {
            Directory.CreateDirectory(directory);
            var path = Path.Combine(directory, name);
            File.WriteAllText(path, content);
            return path;
        }

        private static void WithTempRoot(Action<string> action)
        {
            var root = Path.Combine(
                Path.GetTempPath(),
                "HainanSettlementToolTests",
                Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root);
            try
            {
                action(root);
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
