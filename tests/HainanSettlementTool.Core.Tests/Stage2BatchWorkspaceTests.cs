using System;
using System.IO;
using System.Linq;
using HainanSettlementTool.Core.Services;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace HainanSettlementTool.Core.Tests
{
    [TestClass]
    public sealed class Stage2BatchWorkspaceTests
    {
        [TestMethod]
        public void PublishCopiesNewFilesAndReturnsStrictPathMappings()
        {
            var root = CreateTempRoot();
            try
            {
                var workspace = new Stage2BatchWorkspace(root, "海南", 7);
                var stagedFile = WriteStageFile(workspace, Path.Combine("代理", "A.xlsx"), "new-content");

                var mappings = workspace.Publish();

                var finalFile = Path.Combine(root, "代理", "A.xlsx");
                Assert.IsTrue(workspace.IsPublished);
                Assert.IsFalse(Directory.Exists(workspace.StagingDirectory));
                Assert.AreEqual("new-content", File.ReadAllText(finalFile));
                Assert.AreEqual(finalFile, mappings[stagedFile]);
                Assert.IsTrue(IsStrictlyWithin(root, finalFile));
                Assert.AreEqual(0, CountDirectoriesWithPrefix(root, Stage2BatchWorkspace.BackupDirectoryPrefix));
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void PublishBacksUpExistingFileThenRemovesBackupAfterSuccess()
        {
            var root = CreateTempRoot();
            try
            {
                Directory.CreateDirectory(root);
                var finalFile = Path.Combine(root, "summary.xlsx");
                File.WriteAllText(finalFile, "old-content");

                var workspace = new Stage2BatchWorkspace(root, "重庆", 7);
                WriteStageFile(workspace, "summary.xlsx", "new-content");

                workspace.Publish();

                Assert.AreEqual("new-content", File.ReadAllText(finalFile));
                Assert.AreEqual(0, CountDirectoriesWithPrefix(root, Stage2BatchWorkspace.BackupDirectoryPrefix));
                Assert.AreEqual(0, CountDirectoriesWithPrefix(root, Stage2BatchWorkspace.StagingDirectoryPrefix));
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void LockedPreconditionRunsBeforeFormalMutationAndCanAbortPublish()
        {
            var root = CreateTempRoot();
            try
            {
                Directory.CreateDirectory(root);
                var finalFile = Path.Combine(root, "summary.xlsx");
                File.WriteAllText(finalFile, "old-content");
                var workspace = new Stage2BatchWorkspace(root, "海南", 7);
                var stagedFile = WriteStageFile(workspace, "summary.xlsx", "new-content");
                var callbackCalls = 0;

                var exception = Assert.ThrowsException<Stage2BatchPublishException>(() =>
                    workspace.Publish(() =>
                    {
                        callbackCalls++;
                        Assert.IsTrue(File.Exists(Path.Combine(root, Stage2BatchWorkspace.PublishLockFileName)));
                        Assert.AreEqual("old-content", File.ReadAllText(finalFile));
                        throw new InvalidOperationException("formal state changed");
                    }));

                Assert.AreEqual(1, callbackCalls);
                Assert.IsTrue(exception.RollbackSucceeded);
                StringAssert.Contains(exception.Message, "formal state changed");
                Assert.AreEqual("old-content", File.ReadAllText(finalFile));
                Assert.AreEqual("new-content", File.ReadAllText(stagedFile));
                Assert.IsFalse(File.Exists(Path.Combine(root, Stage2BatchWorkspace.PublishLockFileName)));
                Assert.AreEqual(0, CountDirectoriesWithPrefix(root, Stage2BatchWorkspace.BackupDirectoryPrefix));
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void PublishRollsBackPreviouslyReplacedFilesWhenLaterTargetIsLocked()
        {
            var root = CreateTempRoot();
            try
            {
                Directory.CreateDirectory(root);
                var firstFinal = Path.Combine(root, "01-first.xlsx");
                var lockedFinal = Path.Combine(root, "02-locked.xlsx");
                File.WriteAllText(firstFinal, "first-old");
                File.WriteAllText(lockedFinal, "locked-old");

                var workspace = new Stage2BatchWorkspace(root, "海南", 7);
                var firstStaged = WriteStageFile(workspace, "01-first.xlsx", "first-new");
                var lockedStaged = WriteStageFile(workspace, "02-locked.xlsx", "locked-new");

                Stage2BatchPublishException exception;
                using (var lockStream = new FileStream(lockedFinal, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                {
                    exception = Assert.ThrowsException<Stage2BatchPublishException>(() => workspace.Publish());
                }

                Assert.IsTrue(exception.RollbackSucceeded, string.Join("; ", exception.RollbackErrors));
                Assert.AreEqual("first-old", File.ReadAllText(firstFinal));
                Assert.AreEqual("locked-old", File.ReadAllText(lockedFinal));
                Assert.AreEqual("first-new", File.ReadAllText(firstStaged));
                Assert.AreEqual("locked-new", File.ReadAllText(lockedStaged));
                Assert.AreEqual(0, CountDirectoriesWithPrefix(root, Stage2BatchWorkspace.BackupDirectoryPrefix));
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void PreserveAsFailedKeepsStagedContentAndWritesConspicuousMarker()
        {
            var root = CreateTempRoot();
            try
            {
                var workspace = new Stage2BatchWorkspace(root, "重庆", 7);
                var stagedFile = WriteStageFile(workspace, Path.Combine("代理", "A.xlsx"), "draft-content");

                var failedDirectory = workspace.PreserveAsFailed("汇总表预检未通过\r\n请查看日志");

                Assert.AreEqual(failedDirectory, workspace.FailedDirectory);
                Assert.IsFalse(Directory.Exists(workspace.StagingDirectory));
                StringAssert.StartsWith(Path.GetFileName(failedDirectory), Stage2BatchWorkspace.FailedDirectoryPrefix);
                Assert.IsTrue(IsStrictlyWithin(root, failedDirectory));
                Assert.AreEqual(
                    "draft-content",
                    File.ReadAllText(Path.Combine(failedDirectory, "代理", "A.xlsx")));

                var markerPath = Path.Combine(failedDirectory, Stage2BatchWorkspace.FailedMarkerFileName);
                Assert.IsTrue(File.Exists(markerPath));
                var marker = File.ReadAllText(markerPath);
                StringAssert.Contains(marker, "禁止用于付款或正式结算");
                StringAssert.Contains(marker, "重庆");
                StringAssert.Contains(marker, "7月");
                StringAssert.Contains(marker, "汇总表预检未通过  请查看日志");
                Assert.IsFalse(File.Exists(stagedFile));
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void GetStagingPathRejectsTraversalRootedAndReservedPaths()
        {
            var root = CreateTempRoot();
            try
            {
                var workspace = new Stage2BatchWorkspace(root, "海南", 7);
                var outside = Path.Combine(
                    root,
                    "..",
                    "outside-" + Guid.NewGuid().ToString("N") + ".xlsx");

                Assert.ThrowsException<ArgumentException>(() => workspace.GetStagingPath(Path.Combine("..", "outside.xlsx")));
                Assert.ThrowsException<ArgumentException>(() => workspace.GetStagingPath(Path.GetFullPath(outside)));
                Assert.ThrowsException<ArgumentException>(() => workspace.GetStagingPath(Stage2BatchWorkspace.FailedDirectoryPrefix + "x" + Path.DirectorySeparatorChar + "a.xlsx"));
                Assert.ThrowsException<ArgumentException>(() => workspace.GetStagingPath(Stage2BatchWorkspace.PublishingMarkerPrefix + "x.txt"));
                Assert.ThrowsException<ArgumentException>(() => workspace.GetStagingPath(Stage2BatchWorkspace.PublishLockFileName));

                Assert.IsFalse(File.Exists(Path.GetFullPath(outside)));
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void EmptyWorkspaceCannotBePublishedAndCanBePreservedForInspection()
        {
            var root = CreateTempRoot();
            try
            {
                var workspace = new Stage2BatchWorkspace(root, "海南", 7);

                var exception = Assert.ThrowsException<Stage2BatchPublishException>(() => workspace.Publish());
                Assert.IsTrue(exception.RollbackSucceeded);

                var failedDirectory = workspace.PreserveAsFailed("空批次被拒绝");
                Assert.IsTrue(Directory.Exists(failedDirectory));
                Assert.IsTrue(File.Exists(Path.Combine(
                    failedDirectory,
                    Stage2BatchWorkspace.FailedMarkerFileName)));
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void GetFinalPathMapsOnlyPathsInsideItsStagingWorkspace()
        {
            var root = CreateTempRoot();
            try
            {
                var workspace = new Stage2BatchWorkspace(root, "海南", 7);
                var stagedPath = workspace.GetStagingPath(Path.Combine("代理", "A.xlsx"));

                Assert.AreEqual(Path.Combine(root, "代理", "A.xlsx"), workspace.GetFinalPath(stagedPath));
                Assert.ThrowsException<InvalidOperationException>(() =>
                    workspace.GetFinalPath(Path.Combine(root, "outside.xlsx")));
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void ConstructorRejectsProvinceLabelThatCanChangePathScope()
        {
            var root = CreateTempRoot();
            try
            {
                Assert.ThrowsException<ArgumentException>(() =>
                    new Stage2BatchWorkspace(root, "海南" + Path.DirectorySeparatorChar + "..", 7));
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void ConstructorRejectsVolumeRootAndAbandonedPublishingMarker()
        {
            var root = CreateTempRoot();
            try
            {
                Assert.ThrowsException<ArgumentException>(() =>
                    new Stage2BatchWorkspace(Path.GetPathRoot(root), "海南", 7));

                Directory.CreateDirectory(root);
                var marker = Path.Combine(
                    root,
                    Stage2BatchWorkspace.PublishingMarkerPrefix + "abandoned.txt");
                File.WriteAllText(marker, "do not pay");

                Assert.ThrowsException<InvalidOperationException>(() =>
                    new Stage2BatchWorkspace(root, "海南", 7));
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        [TestMethod]
        public void PublishRechecksAbandonedMarkerAfterWorkspaceWasCreated()
        {
            var root = CreateTempRoot();
            try
            {
                Directory.CreateDirectory(root);
                var finalFile = Path.Combine(root, "summary.xlsx");
                File.WriteAllText(finalFile, "old-content");
                var workspace = new Stage2BatchWorkspace(root, "海南", 7);
                var stagedFile = WriteStageFile(workspace, "summary.xlsx", "new-content");
                var marker = Path.Combine(
                    root,
                    Stage2BatchWorkspace.PublishingMarkerPrefix + "abandoned-after-plan.txt");
                File.WriteAllText(marker, "do not pay");

                var exception = Assert.ThrowsException<Stage2BatchPublishException>(() => workspace.Publish());

                StringAssert.Contains(exception.Message, "发布中标记");
                Assert.AreEqual("old-content", File.ReadAllText(finalFile));
                Assert.AreEqual("new-content", File.ReadAllText(stagedFile));
                Assert.IsTrue(File.Exists(marker));
            }
            finally
            {
                DeleteTempRoot(root);
            }
        }

        private static string WriteStageFile(Stage2BatchWorkspace workspace, string relativePath, string content)
        {
            var path = workspace.GetStagingPath(relativePath);
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            File.WriteAllText(path, content);
            return path;
        }

        private static int CountDirectoriesWithPrefix(string root, string prefix)
        {
            if (!Directory.Exists(root))
            {
                return 0;
            }

            return Directory.GetDirectories(root)
                .Count(path => Path.GetFileName(path).StartsWith(prefix, StringComparison.OrdinalIgnoreCase));
        }

        private static bool IsStrictlyWithin(string root, string candidate)
        {
            var prefix = Path.GetFullPath(root).TrimEnd(Path.DirectorySeparatorChar)
                + Path.DirectorySeparatorChar;
            return Path.GetFullPath(candidate).StartsWith(prefix, StringComparison.OrdinalIgnoreCase);
        }

        private static string CreateTempRoot()
        {
            return Path.Combine(Path.GetTempPath(), "HainanSettlementToolTests", Guid.NewGuid().ToString("N"));
        }

        private static void DeleteTempRoot(string root)
        {
            if (Directory.Exists(root))
            {
                Directory.Delete(root, true);
            }
        }
    }
}
