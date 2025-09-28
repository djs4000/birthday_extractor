// Namespace for the entire Birthday Extractor application.
namespace BirthdayExtractor
{
    // Imports for handling collections, file I/O, cryptography, and JSON serialization.
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Security.Cryptography;
    using System.Text.Json;

    /// <summary>
    /// Stores persisted application preferences, run history, and integration settings.
    /// This class is serialized to a JSON file in the user's local application data folder.
    /// </summary>
    public sealed class AppConfig
    {
        // --- Default settings for the main form ---

        /// <summary>
        /// The default offset (in days from today) for pre-filling the start date of the birthday window.
        /// For example, a value of 28 will make the default start date 4 weeks from now.
        /// </summary>
        public int DefaultStartOffsetDays { get; set; } = 28;

        /// <summary>
        /// The default duration (in days) for the birthday reporting window.
        /// The end date will be calculated as (start date + window days - 1).
        /// </summary>
        public int DefaultWindowDays { get; set; } = 7;

        /// <summary>
        /// The default minimum age a child will be turning to be included in the report.
        /// </summary>
        public int MinAge { get; set; } = 3;

        /// <summary>
        /// The default maximum age a child will be turning to be included in the report.
        /// </summary>
        public int MaxAge { get; set; } = 14;

        /// <summary>
        /// Determines whether the "Write CSV" checkbox is enabled by default on launch.
        /// </summary>
        public bool DefaultWriteCsv { get; set; } = false;

        /// <summary>
        /// Determines whether the "Write XLSX" checkbox is enabled by default on launch.
        /// </summary>
        public bool DefaultWriteXlsx { get; set; } = true;

        // --- Update and integration settings ---

        /// <summary>
        /// If true, the application will check for a newer version on GitHub during startup.
        /// </summary>
        public bool EnableUpdateChecks { get; set; } = true;

        /// <summary>
        /// An optional GitHub Personal Access Token (PAT) for checking updates from private repositories.
        /// TODO: Encrypt this value before saving to disk to enhance security.
        /// </summary>
        public string? GitHubToken { get; set; }

        /// <summary>
        /// Remembers the directory of the last opened CSV file to improve the user experience.
        /// </summary>
        public string? LastCsvFolder { get; set; }

        // --- ERPNext Integration Settings (Webhook and REST API) ---

        /// <summary>
        /// The URL for a future integration to push extracted leads to an ERPNext webhook.
        /// </summary>
        public string? WebhookUrl { get; set; }

        /// <summary>
        /// The optional authorization header value (e.g., a token) for the ERPNext webhook.
        /// TODO: Encrypt this value before saving to disk.
        /// </summary>
        public string? WebhookAuthHeader { get; set; }

        /// <summary>
        /// The base URL for the ERPNext instance (e.g., "https://my-erp.example.com").
        /// </summary>
        public string? ErpNextBaseUrl { get; set; }

        /// <summary>
        /// The API Key for authenticating with the ERPNext REST API.
        /// TODO: Encrypt this value before saving to disk.
        /// </summary>
        public string? ErpNextApiKey { get; set; }

        /// <summary>
        /// The API Secret for authenticating with the ERPNext REST API.
        /// TODO: Encrypt this value before saving to disk.
        /// </summary>
        public string? ErpNextApiSecret { get; set; }

        // --- Phone Number Normalization Settings ---

        /// <summary>
        /// Enables Google's libphonenumber library for more robust phone number validation and normalization.
        /// </summary>
        public bool UseLibPhoneNumber { get; set; } = false;

        /// <summary>
        /// The default region (e.g., "AE" for UAE) to use when parsing ambiguous phone numbers.
        /// </summary>
        public string DefaultRegion { get; set; } = "AE";

        // --- Run History ---

        /// <summary>
        /// A log of previously processed files and date windows to help users avoid duplicate runs.
        /// </summary>
        public List<ProcessedWindow> History { get; set; } = new();

        /// <summary>
        /// A utility function to check if two date ranges overlap.
        /// </summary>
        /// <returns>True if the date ranges intersect.</returns>
        public static bool WindowsOverlap(DateTime s1, DateTime e1, DateTime s2, DateTime e2)
            => s1 <= e2 && s2 <= e1;
    }

    /// <summary>
    /// Represents a record of a single processing run, stored in the configuration history.
    /// </summary>
    public sealed class ProcessedWindow
    {
        /// <summary>
        /// The inclusive start date of the processed birthday window.
        /// </summary>
        public DateTime Start { get; set; }

        /// <summary>
        /// The inclusive end date of the processed birthday window.
        /// </summary>
        public DateTime End { get; set; }

        /// <summary>
        /// The file name of the source CSV.
        /// </summary>
        public string? CsvName { get; set; }

        /// <summary>
        /// The SHA-256 hash of the source CSV file, used to detect if the file has changed.
        /// </summary>
        public string? CsvSha256 { get; set; }

        /// <summary>
        /// The number of birthday records that were extracted in this run.
        /// </summary>
        public int RowCount { get; set; }

        /// <summary>
        /// The timestamp of when the processing run was completed.
        /// </summary>
        public DateTime ProcessedAt { get; set; }
    }

    /// <summary>
    /// A static helper class for loading, saving, and managing the application's configuration.
    /// </summary>
    public static class ConfigStore
    {
        // JSON serializer options for pretty-printing the config file.
        private static readonly JsonSerializerOptions _jsonOpts = new()
        {
            WriteIndented = true
        };

        /// <summary>
        /// Gets the full path to the configuration file, ensuring the directory exists.
        /// The path is typically in `%LOCALAPPDATA%\BirthdayExtractor\config.json`.
        /// </summary>
        public static string GetConfigPath()
        {
            var dir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "BirthdayExtractor");
            Directory.CreateDirectory(dir); // Ensures the directory is created if it doesn't exist.
            return Path.Combine(dir, "config.json");
        }

        /// <summary>
        /// Loads the configuration from disk. If the file doesn't exist or is corrupt,
        /// it creates a new default configuration and backs up the old file if possible.
        /// </summary>
        /// <returns>A loaded or newly created AppConfig instance.</returns>
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
                return cfg ?? new AppConfig(); // Return new config if deserialization results in null.
            }
            catch 
            {
                // TODO: Log the exception to a file for easier debugging of config issues.
                // If the file is corrupted, back it up before creating a new one.
                try
                {
                    var backupPath = Path.ChangeExtension(path, $".bak_{DateTime.Now:yyyyMMddHHmmss}");
                    File.Copy(path, backupPath, overwrite: false);
                }
                catch { /* Ignore errors during backup creation */ }

                // Create and save a fresh configuration.
                var cfg = new AppConfig();
                Save(cfg);
                return cfg;
            }
        }

        /// <summary>
        /// Serializes the provided configuration object to JSON and saves it to disk.
        /// </summary>
        public static void Save(AppConfig cfg)
        {
            var path = GetConfigPath();
            var json = JsonSerializer.Serialize(cfg, _jsonOpts);
            File.WriteAllText(path, json);
        }

        /// <summary>
        /// Computes the SHA-256 hash of a file, returned as an uppercase hex string.
        /// This is used to uniquely identify a CSV file and detect changes.
        /// </summary>
        public static string ComputeSha256(string filePath)
        {
            using var sha = SHA256.Create();
            using var fs = File.OpenRead(filePath);
            var hash = sha.ComputeHash(fs);
            return Convert.ToHexString(hash); // Requires .NET 5+ for this convenient method.
        }
    }
}
