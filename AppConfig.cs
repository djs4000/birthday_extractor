using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text.Json;

namespace BirthdayExtractor
{
    /// <summary>
    /// Stores persisted application preferences and run history.
    /// </summary>
    public sealed class AppConfig
    {
    /// <summary>
    /// Offset (in days from today) used when pre-filling the start date picker.
    /// </summary>
    public int DefaultStartOffsetDays { get; set; } = 28;   // e.g., today + 28
    /// <summary>
    /// How many days to include in the reporting window by default.
    /// </summary>
    public int DefaultWindowDays { get; set; } = 7;     // inclusive window (start .. start+days-1)
    /// <summary>
    /// Smallest turning age that will be considered a candidate.
    /// </summary>
    public int MinAge { get; set; } = 3;
    /// <summary>
    /// Oldest turning age that should remain in the birthday report.
    /// </summary>
    public int MaxAge { get; set; } = 14;


    /// <summary>
    /// Whether CSV export should be ticked on launch.
    /// </summary>
    public bool DefaultWriteCsv { get; set; } = true;
    /// <summary>
    /// Whether XLSX export should be ticked on launch.
    /// </summary>
    public bool DefaultWriteXlsx { get; set; } = true;

    /// <summary>
    /// When true the application will check GitHub for a newer build during startup.
    /// </summary>
    public bool EnableUpdateChecks { get; set; } = true;

    /// <summary>
    /// Optional personal access token used when querying private GitHub releases.
    /// </summary>
    public string? GitHubToken { get; set; }

        // New property
    /// <summary>
    /// Remembers the last CSV directory to improve the browse UX.
    /// </summary>
    public string? LastCsvFolder { get; set; }

        // Future use for ERPNext webhook
    /// <summary>
    /// Future integration endpoint for ERPNext webhook pushes.
    /// </summary>
    public string? WebhookUrl { get; set; }
    /// <summary>
    /// Optional authorization header value for webhook calls.
    /// </summary>
    public string? WebhookAuthHeader { get; set; }

    /// <summary>
    /// Base URL for ERPNext REST API calls.
    /// </summary>
    public string? ErpNextBaseUrl { get; set; }
    /// <summary>
    /// API key portion used for ERPNext token authentication.
    /// </summary>
    public string? ErpNextApiKey { get; set; }
    /// <summary>
    /// API secret portion used for ERPNext token authentication.
    /// </summary>
    public string? ErpNextApiSecret { get; set; }

    /// <summary>
    /// Log of previously processed windows to warn about duplicates.
    /// </summary>
    public List<ProcessedWindow> History { get; set; } = new();

        public static bool WindowsOverlap(DateTime s1, DateTime e1, DateTime s2, DateTime e2)
            => s1 <= e2 && s2 <= e1;
    
        // New: enable libphonenumber for all numbers

    /// <summary>
    /// Enables libphonenumber-based validation when normalizing phones.
    /// </summary>
    public bool UseLibPhoneNumber { get; set; } = false;

        // Default country for parsing ambiguous numbers (e.g. "050...")
    /// <summary>
    /// Region hint passed to libphonenumber for ambiguous numbers.
    /// </summary>
    public string DefaultRegion { get; set; } = "AE";
    }

    /// <summary>
    /// Represents a processed date slice and relevant metadata.
    /// </summary>
    public sealed class ProcessedWindow
    {
    /// <summary>
    /// Inclusive start date that was processed.
    /// </summary>
    public DateTime Start { get; set; }
    /// <summary>
    /// Inclusive end date that was processed.
    /// </summary>
    public DateTime End { get; set; }
    /// <summary>
    /// Name of the source CSV file.
    /// </summary>
    public string? CsvName { get; set; }
    /// <summary>
    /// Hash of the source CSV for quick change detection.
    /// </summary>
    public string? CsvSha256 { get; set; }
    /// <summary>
    /// Number of rows kept for this run.
    /// </summary>
    public int RowCount { get; set; }
    /// <summary>
    /// Timestamp of when the run completed.
    /// </summary>
    public DateTime ProcessedAt { get; set; }

    }

    /// <summary>
    /// Helper for serializing configuration to disk in the user profile.
    /// </summary>
    public static class ConfigStore
    {
        private static readonly JsonSerializerOptions _jsonOpts = new()
        {
            WriteIndented = true
        };

        /// <summary>
        /// Resolves the user-specific config path, ensuring the folder exists.
        /// </summary>
        public static string GetConfigPath()
        {
            var dir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "BirthdayExtractor");
            Directory.CreateDirectory(dir);
            return Path.Combine(dir, "config.json");
        }

        /// <summary>
        /// Reads configuration from disk, creating defaults (and backups) when required.
        /// </summary>
        public static AppConfig LoadOrCreate()
        {
            var path = GetConfigPath();
            if (!File.Exists(path))
            {
                var cfg = new AppConfig();
                Save(cfg);
                return cfg;
            }

            try
            {
                var json = File.ReadAllText(path);
                var cfg = JsonSerializer.Deserialize<AppConfig>(json, _jsonOpts);
                return cfg ?? new AppConfig();
            }
            catch
            {
                // If file is corrupted, back it up and recreate
                try
                {
                    var backup = Path.ChangeExtension(path, $".bak_{DateTime.Now:yyyyMMddHHmmss}");
                    File.Copy(path, backup, overwrite: false);
                }
                catch { /* ignore */ }

                var cfg = new AppConfig();
                Save(cfg);
                return cfg;
            }
        }

        /// <summary>
        /// Persists the supplied configuration to disk.
        /// </summary>
        public static void Save(AppConfig cfg)
        {
            var path = GetConfigPath();
            var json = JsonSerializer.Serialize(cfg, _jsonOpts);
            File.WriteAllText(path, json);
        }

        /// <summary>
        /// Computes an uppercase SHA-256 hash for the provided file path.
        /// </summary>
        public static string ComputeSha256(string filePath)
        {
            using var sha = SHA256.Create();
            using var fs = File.OpenRead(filePath);
            var hash = sha.ComputeHash(fs);
            return Convert.ToHexString(hash); // .NET 5+ uppercase hex
        }

    }
    
}