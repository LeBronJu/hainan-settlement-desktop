using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

namespace HainanSettlementTool.Core.Services
{
    /// <summary>
    /// Provides an isolated workspace for one Stage 2 batch and publishes its
    /// files to the requested output root only after every file is ready.
    /// </summary>
    public sealed class Stage2BatchWorkspace
    {
        public const string StagingDirectoryPrefix = "【处理中-阶段二】";
        public const string BackupDirectoryPrefix = "【阶段二发布备份】";
        public const string FailedDirectoryPrefix = "【未完成-禁止付款】";
        public const string FailedMarkerFileName = "【未完成-禁止付款】说明.txt";
        public const string PublishingMarkerPrefix = "【发布中-禁止付款】";
        public const string PublishLockFileName = "【阶段二发布锁】.lock";

        private readonly string _provinceLabel;
        private readonly int _month;
        private WorkspaceState _state;
        private string _lastBackupDirectory;

        public Stage2BatchWorkspace(string outputDirectory, string provinceLabel, int month)
        {
            if (month < 1 || month > 12)
            {
                throw new ArgumentOutOfRangeException(nameof(month), "结算月份必须在1到12之间。");
            }

            _provinceLabel = ValidateLabel(provinceLabel);
            _month = month;
            OutputDirectory = NormalizeDirectory(outputDirectory, "阶段二正式输出根目录");
            Directory.CreateDirectory(OutputDirectory);
            if (IsReparsePoint(OutputDirectory))
            {
                throw new IOException("阶段二正式输出根目录不能是重解析目录：" + OutputDirectory);
            }

            EnsureNoPublishingMarker();

            StagingDirectory = CreateUniqueChildDirectory(OutputDirectory, StagingDirectoryPrefix);
            EnsureStrictlyWithin(OutputDirectory, StagingDirectory, "阶段二 staging 目录");
            Directory.CreateDirectory(StagingDirectory);
            _state = WorkspaceState.Ready;
        }

        public string OutputDirectory { get; }

        public string StagingDirectory { get; }

        public string FailedDirectory { get; private set; }

        public bool IsPublished
        {
            get { return _state == WorkspaceState.Published; }
        }

        /// <summary>
        /// Resolves a caller-supplied relative path below the staging root.
        /// Rooted paths, traversal, and reserved workspace directory names are rejected.
        /// </summary>
        public string GetStagingPath(string relativePath)
        {
            EnsureStateAllowsWriting();
            ValidateRelativePath(relativePath);
            var candidate = Path.GetFullPath(Path.Combine(StagingDirectory, relativePath));
            EnsureStrictlyWithin(StagingDirectory, candidate, "staging 输出路径");
            EnsureStrictlyWithin(OutputDirectory, candidate, "staging 输出路径");
            return candidate;
        }

        /// <summary>
        /// Resolves the formal destination for a path inside this staging workspace.
        /// This is useful when reports must contain final paths before the batch is
        /// published. The method does not publish or create the destination.
        /// </summary>
        public string GetFinalPath(string stagingPath)
        {
            if (string.IsNullOrWhiteSpace(stagingPath))
            {
                throw new ArgumentException("请提供 staging 路径。", nameof(stagingPath));
            }

            var fullStagingPath = Path.GetFullPath(stagingPath);
            var relativePath = GetRelativePathStrict(StagingDirectory, fullStagingPath);
            ValidateRelativePath(relativePath);
            var finalPath = Path.GetFullPath(Path.Combine(OutputDirectory, relativePath));
            EnsureStrictlyWithin(OutputDirectory, finalPath, "阶段二正式输出路径");
            EnsureNotWithin(StagingDirectory, finalPath, "阶段二正式输出路径");
            return finalPath;
        }

        /// <summary>
        /// Publishes every staging file below the formal output root while preserving
        /// relative paths. Existing files are moved to a private rollback directory
        /// before replacement. A failed publish makes a best-effort rollback.
        /// </summary>
        public IReadOnlyDictionary<string, string> Publish()
        {
            return Publish(null);
        }

        /// <summary>
        /// Publishes the batch while holding the output-root publish lock. The optional
        /// precondition runs after the lock is acquired and before any formal output is
        /// moved or copied, closing the gap between final state validation and publish.
        /// </summary>
        public IReadOnlyDictionary<string, string> Publish(Action lockedPrecondition)
        {
            if (_state != WorkspaceState.Ready)
            {
                throw new InvalidOperationException("当前阶段二工作区状态不允许发布。");
            }

            if (!Directory.Exists(StagingDirectory))
            {
                throw new DirectoryNotFoundException("阶段二 staging 目录不存在：" + StagingDirectory);
            }

            _state = WorkspaceState.Publishing;
            var backupDirectory = CreateUniqueChildDirectory(OutputDirectory, BackupDirectoryPrefix);
            _lastBackupDirectory = backupDirectory;
            EnsureStrictlyWithin(OutputDirectory, backupDirectory, "阶段二发布备份目录");

            var entries = new List<PublishEntry>();
            var mapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var applied = new List<PublishEntry>();
            var createdFinalDirectories = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            FileStream publishLock = null;
            string publishingMarker = null;
            try
            {
                publishLock = AcquirePublishLock();
                EnsureNoPublishingMarker();
                lockedPrecondition?.Invoke();
                publishingMarker = CreatePublishingMarker();
                entries = BuildPublishEntries(backupDirectory);
                if (entries.Count == 0)
                {
                    throw new InvalidOperationException("阶段二 staging 工作区没有任何待发布文件，已拒绝空批次发布。");
                }

                foreach (var entry in entries)
                {
                    mapping.Add(entry.StagingPath, entry.FinalPath);
                }

                Directory.CreateDirectory(backupDirectory);

                foreach (var entry in entries)
                {
                    EnsureDirectoryAndTrack(
                        Path.GetDirectoryName(entry.FinalPath),
                        OutputDirectory,
                        createdFinalDirectories);
                    EnsureExistingPathSegmentsAreNotReparsePoints(OutputDirectory, entry.FinalPath);

                    applied.Add(entry);
                    if (File.Exists(entry.FinalPath))
                    {
                        Directory.CreateDirectory(Path.GetDirectoryName(entry.BackupPath));
                        File.Move(entry.FinalPath, entry.BackupPath);
                        entry.OriginalWasBackedUp = true;
                    }

                    entry.CopyWasStarted = true;
                    File.Copy(entry.StagingPath, entry.FinalPath, false);
                    entry.WasPublished = true;
                }

                Directory.Delete(StagingDirectory, true);
                File.Delete(publishingMarker);
                publishingMarker = null;
                TryDeleteDirectory(backupDirectory);
                _lastBackupDirectory = Directory.Exists(backupDirectory) ? backupDirectory : null;
                _state = WorkspaceState.Published;
                return new ReadOnlyDictionary<string, string>(mapping);
            }
            catch (Exception ex)
            {
                var rollbackErrors = RollBack(applied, createdFinalDirectories, backupDirectory);
                if (rollbackErrors.Count == 0)
                {
                    TryDeleteFile(publishingMarker);
                    publishingMarker = null;
                }

                _state = WorkspaceState.Failed;
                throw new Stage2BatchPublishException(
                    "阶段二批次发布失败，已尽力回滚正式输出目录。原因：" + ShortReason(ex.Message),
                    ex,
                    rollbackErrors.Count == 0,
                    Directory.Exists(backupDirectory) ? backupDirectory : null,
                    rollbackErrors);
            }
            finally
            {
                if (publishLock != null)
                {
                    publishLock.Dispose();
                }

                TryDeleteFile(Path.Combine(OutputDirectory, PublishLockFileName));
            }
        }

        /// <summary>
        /// Renames a non-published staging workspace to a conspicuous failure directory
        /// and writes a short marker without deleting any staged content.
        /// </summary>
        public string PreserveAsFailed(string reason)
        {
            if (_state == WorkspaceState.Published || _state == WorkspaceState.Preserved)
            {
                throw new InvalidOperationException("当前阶段二工作区状态不允许标记为未完成。");
            }

            if (_state == WorkspaceState.Publishing)
            {
                throw new InvalidOperationException("阶段二工作区正在发布，不能同时标记失败。");
            }

            if (!Directory.Exists(StagingDirectory))
            {
                throw new DirectoryNotFoundException("需保留的 staging 目录不存在：" + StagingDirectory);
            }

            var failedDirectory = CreateUniqueChildDirectory(OutputDirectory, FailedDirectoryPrefix);
            EnsureStrictlyWithin(OutputDirectory, failedDirectory, "阶段二未完成目录");
            Directory.Move(StagingDirectory, failedDirectory);
            FailedDirectory = failedDirectory;
            _state = WorkspaceState.Preserved;

            var markerPath = Path.Combine(failedDirectory, FailedMarkerFileName);
            EnsureStrictlyWithin(failedDirectory, markerPath, "阶段二未完成标记文件");
            var marker = new StringBuilder();
            marker.AppendLine("此目录为阶段二未完成工作区，禁止用于付款或正式结算。");
            marker.AppendLine("省份：" + _provinceLabel);
            marker.AppendLine("月份：" + _month.ToString(CultureInfo.InvariantCulture) + "月");
            marker.AppendLine("原因：" + ShortReason(reason));
            if (!string.IsNullOrWhiteSpace(_lastBackupDirectory) && Directory.Exists(_lastBackupDirectory))
            {
                marker.AppendLine("回滚备份：" + _lastBackupDirectory);
            }

            File.WriteAllText(markerPath, marker.ToString(), new UTF8Encoding(true));
            return failedDirectory;
        }

        private List<PublishEntry> BuildPublishEntries(string backupDirectory)
        {
            var stagingFiles = new List<string>();
            CollectWorkspaceFiles(StagingDirectory, stagingFiles);

            var entries = new List<PublishEntry>();
            var finalPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var source in stagingFiles)
            {
                var relativePath = GetRelativePathStrict(StagingDirectory, source);
                ValidateRelativePath(relativePath);

                var finalPath = Path.GetFullPath(Path.Combine(OutputDirectory, relativePath));
                var backupPath = Path.GetFullPath(Path.Combine(backupDirectory, relativePath));
                EnsureStrictlyWithin(OutputDirectory, finalPath, "阶段二正式输出文件");
                EnsureStrictlyWithin(backupDirectory, backupPath, "阶段二发布备份文件");
                EnsureNotWithin(StagingDirectory, finalPath, "阶段二正式输出文件");

                if (!finalPaths.Add(finalPath))
                {
                    throw new InvalidOperationException("阶段二 staging 中有多个文件指向同一正式输出路径：" + finalPath);
                }

                if (Directory.Exists(finalPath))
                {
                    throw new IOException("阶段二正式输出文件与已有目录冲突：" + finalPath);
                }

                if (File.Exists(finalPath) && IsReparsePoint(finalPath))
                {
                    throw new IOException("阶段二正式输出目标是重解析文件，已拒绝：" + finalPath);
                }

                EnsureExistingPathSegmentsAreNotReparsePoints(OutputDirectory, finalPath);
                entries.Add(new PublishEntry(source, finalPath, backupPath, relativePath));
            }

            return entries
                .OrderBy(entry => entry.RelativePath, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static void CollectWorkspaceFiles(string directory, IList<string> files)
        {
            if (IsReparsePoint(directory))
            {
                throw new IOException("staging 目录不能包含重解析目录：" + directory);
            }

            foreach (var file in Directory.GetFiles(directory, "*", SearchOption.TopDirectoryOnly))
            {
                if (IsReparsePoint(file))
                {
                    throw new IOException("staging 目录不能包含重解析文件：" + file);
                }

                files.Add(Path.GetFullPath(file));
            }

            foreach (var child in Directory.GetDirectories(directory, "*", SearchOption.TopDirectoryOnly))
            {
                if (IsReparsePoint(child))
                {
                    throw new IOException("staging 目录不能包含重解析目录：" + child);
                }

                CollectWorkspaceFiles(child, files);
            }
        }

        private List<string> RollBack(
            IList<PublishEntry> applied,
            ISet<string> createdFinalDirectories,
            string backupDirectory)
        {
            var errors = new List<string>();
            for (var index = applied.Count - 1; index >= 0; index--)
            {
                var entry = applied[index];
                if ((entry.CopyWasStarted || entry.WasPublished) && File.Exists(entry.FinalPath))
                {
                    if (!File.Exists(entry.StagingPath))
                    {
                        try
                        {
                            Directory.CreateDirectory(Path.GetDirectoryName(entry.StagingPath));
                            File.Copy(entry.FinalPath, entry.StagingPath, false);
                        }
                        catch (Exception rollbackException)
                        {
                            errors.Add(
                                entry.RelativePath
                                + "（恢复 staging）："
                                + rollbackException.Message);
                        }
                    }

                    try
                    {
                        File.Delete(entry.FinalPath);
                    }
                    catch (Exception rollbackException)
                    {
                        errors.Add(
                            entry.RelativePath
                            + "（移除未完成正式文件）："
                            + rollbackException.Message);
                    }
                }

                if (!File.Exists(entry.StagingPath) && !File.Exists(entry.FinalPath))
                {
                    errors.Add(entry.RelativePath + "（staging 文件无法恢复）：文件已不存在。");
                }

                if (!entry.OriginalWasBackedUp)
                {
                    continue;
                }

                if (!File.Exists(entry.BackupPath))
                {
                    errors.Add(entry.RelativePath + "（恢复原文件）：回滚备份已不存在。");
                    continue;
                }

                if (File.Exists(entry.FinalPath) || Directory.Exists(entry.FinalPath))
                {
                    errors.Add(entry.RelativePath + "（恢复原文件）：正式目标仍被占用，原文件保留在回滚备份中。");
                    continue;
                }

                try
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(entry.FinalPath));
                    File.Move(entry.BackupPath, entry.FinalPath);
                }
                catch (Exception rollbackException)
                {
                    errors.Add(
                        entry.RelativePath
                        + "（恢复原文件）："
                        + rollbackException.Message);
                }
            }

            foreach (var directory in createdFinalDirectories
                .OrderByDescending(path => path.Length)
                .ToList())
            {
                try
                {
                    if (Directory.Exists(directory) && !Directory.EnumerateFileSystemEntries(directory).Any())
                    {
                        Directory.Delete(directory, false);
                    }
                }
                catch (Exception rollbackException)
                {
                    errors.Add("清理目录 " + directory + "：" + rollbackException.Message);
                }
            }

            if (errors.Count == 0)
            {
                TryDeleteDirectory(backupDirectory);
                if (!Directory.Exists(backupDirectory))
                {
                    _lastBackupDirectory = null;
                }
            }

            return errors;
        }

        private static void EnsureDirectoryAndTrack(string directory, string root, ISet<string> createdDirectories)
        {
            var missing = new Stack<string>();
            var current = Path.GetFullPath(directory);
            if (!string.Equals(Path.GetFullPath(root), current, StringComparison.OrdinalIgnoreCase))
            {
                EnsureStrictlyWithin(root, current, "阶段二正式输出目录");
            }

            while (!Directory.Exists(current))
            {
                missing.Push(current);
                var parent = Path.GetDirectoryName(current);
                if (string.IsNullOrWhiteSpace(parent))
                {
                    throw new IOException("无法在正式输出根目录内创建目标目录：" + directory);
                }

                current = parent;
            }

            while (missing.Count > 0)
            {
                var path = missing.Pop();
                Directory.CreateDirectory(path);
                createdDirectories.Add(path);
            }
        }

        private static void EnsureExistingPathSegmentsAreNotReparsePoints(string root, string candidate)
        {
            var relative = GetRelativePathStrict(root, candidate);
            var segments = relative.Split(new[] { Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar }, StringSplitOptions.RemoveEmptyEntries);
            var current = root;
            for (var index = 0; index < segments.Length - 1; index++)
            {
                current = Path.Combine(current, segments[index]);
                if (Directory.Exists(current) && IsReparsePoint(current))
                {
                    throw new IOException("阶段二输出路径包含重解析目录，已拒绝：" + current);
                }
            }
        }

        private void EnsureStateAllowsWriting()
        {
            if (_state != WorkspaceState.Ready && _state != WorkspaceState.Failed)
            {
                throw new InvalidOperationException("当前阶段二工作区状态不允许写入 staging。");
            }
        }

        private static string NormalizeDirectory(string path, string label)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                throw new ArgumentException("请提供" + label + "。", nameof(path));
            }

            var fullPath = Path.GetFullPath(path.Trim());
            var root = Path.GetPathRoot(fullPath);
            if (string.Equals(
                fullPath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar),
                root == null ? string.Empty : root.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar),
                StringComparison.OrdinalIgnoreCase))
            {
                throw new ArgumentException(label + "不能直接使用卷根目录。", nameof(path));
            }

            return fullPath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        }

        private static string ValidateLabel(string value)
        {
            var label = TextUtil.S(value);
            if (string.IsNullOrWhiteSpace(label))
            {
                throw new ArgumentException("请提供阶段二省份标签。", nameof(value));
            }

            if (label.Length > 40
                || label == "."
                || label == ".."
                || label.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0
                || label.IndexOf(Path.DirectorySeparatorChar) >= 0
                || label.IndexOf(Path.AltDirectorySeparatorChar) >= 0)
            {
                throw new ArgumentException("阶段二省份标签不能包含路径或非法文件名字符。", nameof(value));
            }

            return label;
        }

        private string CreateUniqueChildDirectory(string root, string prefix)
        {
            var name = prefix
                + _provinceLabel
                + "-"
                + _month.ToString(CultureInfo.InvariantCulture)
                + "月-"
                + DateTime.Now.ToString("yyyyMMdd-HHmmssfff", CultureInfo.InvariantCulture)
                + "-"
                + Guid.NewGuid().ToString("N");
            var candidate = Path.GetFullPath(Path.Combine(root, name));
            EnsureStrictlyWithin(root, candidate, "阶段二工作区子目录");
            return candidate;
        }

        private static void ValidateRelativePath(string relativePath)
        {
            if (string.IsNullOrWhiteSpace(relativePath))
            {
                throw new ArgumentException("阶段二工作区路径必须是非空相对路径。", nameof(relativePath));
            }

            if (Path.IsPathRooted(relativePath))
            {
                throw new ArgumentException("阶段二工作区路径不能是绝对路径。", nameof(relativePath));
            }

            var segments = relativePath.Split(
                new[] { Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar },
                StringSplitOptions.RemoveEmptyEntries);
            if (segments.Length == 0 || segments.Any(segment => segment == "." || segment == ".."))
            {
                throw new ArgumentException("阶段二工作区路径不能越界。", nameof(relativePath));
            }

            if (segments.Any(segment => segment.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0))
            {
                throw new ArgumentException("阶段二工作区路径包含非法文件名字符。", nameof(relativePath));
            }

            var first = segments[0];
            if (first.StartsWith(StagingDirectoryPrefix, StringComparison.OrdinalIgnoreCase)
                || first.StartsWith(BackupDirectoryPrefix, StringComparison.OrdinalIgnoreCase)
                || first.StartsWith(FailedDirectoryPrefix, StringComparison.OrdinalIgnoreCase)
                || first.StartsWith(PublishingMarkerPrefix, StringComparison.OrdinalIgnoreCase)
                || string.Equals(first, PublishLockFileName, StringComparison.OrdinalIgnoreCase))
            {
                throw new ArgumentException("阶段二工作区路径不能使用系统保留目录名。", nameof(relativePath));
            }
        }

        private static string GetRelativePathStrict(string root, string candidate)
        {
            var fullRoot = Path.GetFullPath(root);
            var fullCandidate = Path.GetFullPath(candidate);
            EnsureStrictlyWithin(fullRoot, fullCandidate, "阶段二相对路径");
            return fullCandidate.Substring(AppendDirectorySeparator(fullRoot).Length);
        }

        private static void EnsureStrictlyWithin(string root, string candidate, string label)
        {
            var fullRoot = Path.GetFullPath(root);
            var fullCandidate = Path.GetFullPath(candidate);
            var rootPrefix = AppendDirectorySeparator(fullRoot);
            if (!fullCandidate.StartsWith(rootPrefix, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(label + "必须严格位于明确的根目录内：" + fullRoot);
            }
        }

        private static void EnsureNotWithin(string forbiddenRoot, string candidate, string label)
        {
            var fullForbiddenRoot = Path.GetFullPath(forbiddenRoot);
            var fullCandidate = Path.GetFullPath(candidate);
            if (string.Equals(fullForbiddenRoot, fullCandidate, StringComparison.OrdinalIgnoreCase)
                || fullCandidate.StartsWith(AppendDirectorySeparator(fullForbiddenRoot), StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(label + "不能回写到 staging 目录。");
            }
        }

        private static string AppendDirectorySeparator(string path)
        {
            return path.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal)
                || path.EndsWith(Path.AltDirectorySeparatorChar.ToString(), StringComparison.Ordinal)
                ? path
                : path + Path.DirectorySeparatorChar;
        }

        private static bool IsReparsePoint(string path)
        {
            return (File.GetAttributes(path) & FileAttributes.ReparsePoint) != 0;
        }

        private static void TryDeleteDirectory(string path)
        {
            try
            {
                if (Directory.Exists(path))
                {
                    Directory.Delete(path, true);
                }
            }
            catch
            {
                // A published batch is still valid if a private backup cannot be
                // cleaned immediately. The backup path remains conspicuous.
            }
        }

        private FileStream AcquirePublishLock()
        {
            var path = Path.Combine(OutputDirectory, PublishLockFileName);
            EnsureStrictlyWithin(OutputDirectory, path, "阶段二发布锁");
            try
            {
                return new FileStream(path, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
            }
            catch (Exception ex)
            {
                throw new IOException("另一个阶段二批次正在发布到同一输出目录，本批已停止。", ex);
            }
        }

        private void EnsureNoPublishingMarker()
        {
            var abandonedMarker = Directory.GetFiles(
                    OutputDirectory,
                    PublishingMarkerPrefix + "*",
                    SearchOption.TopDirectoryOnly)
                .FirstOrDefault();
            if (!string.IsNullOrWhiteSpace(abandonedMarker))
            {
                throw new InvalidOperationException(
                    "阶段二输出目录存在未清理的发布中标记，上一批可能在发布时异常中断。请先人工核查正式文件和回滚备份："
                    + abandonedMarker);
            }
        }

        private string CreatePublishingMarker()
        {
            var name = PublishingMarkerPrefix
                + _provinceLabel
                + "-"
                + _month.ToString(CultureInfo.InvariantCulture)
                + "月-"
                + Guid.NewGuid().ToString("N")
                + ".txt";
            var path = Path.Combine(OutputDirectory, name);
            EnsureStrictlyWithin(OutputDirectory, path, "阶段二发布中标记");
            File.WriteAllText(
                path,
                "阶段二批次正在发布，在此标记消失前禁止将目录中文件用于付款或正式结算。",
                new UTF8Encoding(true));
            return path;
        }

        private static void TryDeleteFile(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return;
            }

            try
            {
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
            }
            catch
            {
                // A retained marker is intentionally safer than hiding a possibly
                // interrupted publish. The next workspace will refuse to continue.
            }
        }

        private static string ShortReason(string reason)
        {
            var text = TextUtil.S(reason);
            if (string.IsNullOrWhiteSpace(text))
            {
                return "未提供详细原因，请查看程序日志。";
            }

            text = text.Replace("\r", " ").Replace("\n", " ");
            return text.Length <= 240 ? text : text.Substring(0, 240) + "...";
        }

        private enum WorkspaceState
        {
            Ready,
            Publishing,
            Failed,
            Preserved,
            Published
        }

        private sealed class PublishEntry
        {
            public PublishEntry(string stagingPath, string finalPath, string backupPath, string relativePath)
            {
                StagingPath = stagingPath;
                FinalPath = finalPath;
                BackupPath = backupPath;
                RelativePath = relativePath;
            }

            public string StagingPath { get; }

            public string FinalPath { get; }

            public string BackupPath { get; }

            public string RelativePath { get; }

            public bool OriginalWasBackedUp { get; set; }

            public bool CopyWasStarted { get; set; }

            public bool WasPublished { get; set; }
        }
    }

    public sealed class Stage2BatchPublishException : IOException
    {
        public Stage2BatchPublishException(
            string message,
            Exception innerException,
            bool rollbackSucceeded,
            string retainedBackupDirectory,
            IList<string> rollbackErrors)
            : base(message, innerException)
        {
            RollbackSucceeded = rollbackSucceeded;
            RetainedBackupDirectory = retainedBackupDirectory;
            RollbackErrors = new ReadOnlyCollection<string>(
                rollbackErrors == null ? new List<string>() : rollbackErrors.ToList());
        }

        public bool RollbackSucceeded { get; }

        public string RetainedBackupDirectory { get; }

        public IReadOnlyList<string> RollbackErrors { get; }
    }
}
