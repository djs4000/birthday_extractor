using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Windows.Forms;

namespace BirthdayExtractor
{
    /// <summary>
    /// Handles scenarios where an older version of the application downloads a newer executable
    /// to a temporary directory and simply launches it. The newer build (running from temp) needs
    /// to copy itself back over the original install location before continuing.
    /// </summary>
    internal static class SelfUpdateCoordinator
    {
        /// <summary>
        /// Performs self-replacement when the app is executed from a temporary directory and persists
        /// the install location when running from the normal installation path.
        /// </summary>
        /// <returns>True when the current process should exit because a replacement and relaunch occurred.</returns>
        public static bool TryHandleSelfReplacement()
        {
            try
            {
                var currentExe = Application.ExecutablePath;
                if (string.IsNullOrWhiteSpace(currentExe) || !File.Exists(currentExe))
                {
                    return false;
                }

                var currentPath = Path.GetFullPath(currentExe);
                var tempRoot = Path.GetFullPath(Path.GetTempPath());

                if (!currentPath.StartsWith(tempRoot, StringComparison.OrdinalIgnoreCase))
                {
                    PersistInstallLocation(currentPath);
                    return false;
                }

                var candidateTargets = GatherCandidateTargets(currentPath);
                foreach (var target in candidateTargets)
                {
                    if (TryReplaceExecutable(currentPath, target))
                    {
                        return true;
                    }
                }

                LogRouter.LogMessage("Self-update: running from temporary folder but no valid installation target was found.");
            }
            catch (Exception ex)
            {
                LogRouter.LogException(ex, "Self-update replacement failed");
            }

            return false;
        }

        private static IReadOnlyCollection<string> GatherCandidateTargets(string currentPath)
        {
            var candidates = new List<string>();
            var fileName = Path.GetFileName(currentPath);

            var envTarget = Environment.GetEnvironmentVariable("BIRTHDAY_EXTRACTOR_TARGET_EXE");
            if (!string.IsNullOrWhiteSpace(envTarget))
            {
                TryAddCandidate(candidates, envTarget);
            }

            var workingDirectory = Environment.CurrentDirectory;
            if (!string.IsNullOrWhiteSpace(workingDirectory))
            {
                TryAddCandidate(candidates, Path.Combine(workingDirectory, fileName));
            }

            try
            {
                var cfg = ConfigStore.LoadOrCreate();
                if (!string.IsNullOrWhiteSpace(cfg.LastInstalledExecutable))
                {
                    TryAddCandidate(candidates, cfg.LastInstalledExecutable!);
                }
            }
            catch (Exception ex)
            {
                LogRouter.LogException(ex, "Self-update: failed to read configuration while gathering targets");
            }

            return candidates;
        }

        private static void TryAddCandidate(List<string> candidates, string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return;
            }

            try
            {
                var normalized = Path.GetFullPath(path);
                foreach (var existing in candidates)
                {
                    if (string.Equals(existing, normalized, StringComparison.OrdinalIgnoreCase))
                    {
                        return;
                    }
                }

                candidates.Add(normalized);
            }
            catch (Exception ex)
            {
                LogRouter.LogException(ex, "Self-update: failed to normalize candidate path");
            }
        }

        private static void PersistInstallLocation(string currentPath)
        {
            try
            {
                var cfg = ConfigStore.LoadOrCreate();
                if (!string.Equals(cfg.LastInstalledExecutable, currentPath, StringComparison.OrdinalIgnoreCase))
                {
                    cfg.LastInstalledExecutable = currentPath;
                    ConfigStore.Save(cfg);
                }
            }
            catch (Exception ex)
            {
                LogRouter.LogException(ex, "Self-update: failed to persist install location");
            }
        }

        private static bool TryReplaceExecutable(string sourceExe, string targetExe)
        {
            if (string.IsNullOrWhiteSpace(targetExe))
            {
                return false;
            }

            try
            {
                var normalizedTarget = Path.GetFullPath(targetExe);
                if (string.Equals(sourceExe, normalizedTarget, StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }

                if (!File.Exists(normalizedTarget))
                {
                    return false;
                }

                var targetDir = Path.GetDirectoryName(normalizedTarget);
                if (string.IsNullOrWhiteSpace(targetDir))
                {
                    return false;
                }

                Directory.CreateDirectory(targetDir);

                var stagedCopy = Path.Combine(Path.GetTempPath(), $"be_selfupdate_{Guid.NewGuid():N}.tmp");
                File.Copy(sourceExe, stagedCopy, overwrite: true);

                const int maxAttempts = 20;
                for (var attempt = 0; attempt < maxAttempts && File.Exists(normalizedTarget); attempt++)
                {
                    try
                    {
                        File.Delete(normalizedTarget);
                    }
                    catch (IOException)
                    {
                        Thread.Sleep(250);
                    }
                    catch (UnauthorizedAccessException)
                    {
                        Thread.Sleep(250);
                    }
                }

                if (File.Exists(normalizedTarget))
                {
                    TryDelete(stagedCopy);
                    return false;
                }

                File.Move(stagedCopy, normalizedTarget);
                PersistInstallLocation(normalizedTarget);
                LogRouter.LogMessage($"Self-update: replaced executable at {normalizedTarget}.");

                try
                {
                    Process.Start(new ProcessStartInfo(normalizedTarget)
                    {
                        UseShellExecute = true
                    });
                }
                catch (Exception relaunchEx)
                {
                    LogRouter.LogException(relaunchEx, "Self-update: failed to relaunch application after replacement");
                }

                Environment.Exit(0);
                return true; // unreachable, but satisfies static analysis.
            }
            catch (Exception ex)
            {
                LogRouter.LogException(ex, "Self-update: failed to copy executable");
            }

            return false;
        }

        private static void TryDelete(string path)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(path) && File.Exists(path))
                {
                    File.Delete(path);
                }
            }
            catch
            {
                // Swallow cleanup failures.
            }
        }
    }
}

