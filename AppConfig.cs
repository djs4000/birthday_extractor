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
        private static string DigitsOnly(string s)
        {
            var sb = new System.Text.StringBuilder(s.Length);
            foreach (var ch in s)
                if (ch >= '0' && ch <= '9') sb.Append(ch);
            return sb.ToString();
        }

        /// <summary>
        /// Try to normalize a UAE mobile to E.164 (+9715########). Returns true if recognized as UAE mobile.
        /// Accepts inputs like +9715..., 009715..., 9715..., 05..., 5... and strips punctuation/spaces.
        /// </summary>
        /// <summary>
        /// Attempts to normalize known UAE mobile formats into E.164 representation.
        /// </summary>
        private static bool TryNormalizeUaeMobile(string? input, out string normalized)
        {
            normalized = string.Empty;
            if (string.IsNullOrWhiteSpace(input)) return false;

            var raw = input.Trim();
            bool hadPlus = raw.StartsWith("+");
            var digits = DigitsOnly(raw);

            // Handle 00 prefix (international)
            if (digits.StartsWith("00")) digits = digits.Substring(2);

            // If starts with 971 and 12 digits total: expect 971 5 ########
            if (digits.Length == 12 && digits.StartsWith("971") && digits[3] == '5')
            {
                normalized = "+971" + digits.Substring(3); // +9715########
                return true;
            }

            // If starts with 971 but not 12 digits yet (e.g. formatting remnants)
            if (digits.StartsWith("971"))
            {
                var tail = digits.Substring(3);
                if (tail.Length == 9 && tail.StartsWith("5"))
                {
                    normalized = "+971" + tail;
                    return true;
                }
            }

            // If local 9-digit mobile starting with 5x
            if (digits.Length == 9 && digits[0] == '5' && IsUaeMobilePrefix(digits.Substring(0, 2)))
            {
                normalized = "+971" + digits; // assume UAE
                return true;
            }

            // If starts with 0 then 5x and 10 digits (e.g., 05########)
            if (digits.Length == 10 && digits[0] == '0' && digits[1] == '5' && IsUaeMobilePrefix(digits.Substring(1, 2)))
            {
                normalized = "+971" + digits.Substring(1); // drop leading 0
                return true;
            }

            // If someone typed only 8 digits (missing leading 5?), don’t guess.
            return false;
        }

        /// <summary>
        /// Returns true when the two-digit prefix matches a known UAE mobile range.
        /// </summary>
        private static bool IsUaeMobilePrefix(string twoDigits)
        {
            // common mobile prefixes in UAE: 50/52/53/54/55/56/57/58 (some are MVNO/operator-specific)
            return twoDigits is "50" or "52" or "53" or "54" or "55" or "56" or "57" or "58";
        }

        /// <summary>
        /// Normalization for matching across all phones:
        /// - If UAE mobile pattern recognized -> E.164 (+9715########)
        /// - Else: return canonical digits (with leading '+' if present) for stable matching, no validity asserted.
        /// </summary>
        /// <summary>
        /// Produces a simplified phone token that aids in matching across records.
        /// </summary>
        private static string NormalizePhoneForMatching(string? input)
        {
            if (string.IsNullOrWhiteSpace(input)) return string.Empty;

            if (TryNormalizeUaeMobile(input, out var uae)) return uae;

            // Generic: keep a '+' if it was at the start, but strip everything else non-digit
            string trimmed = input.Trim();
            bool leadingPlus = trimmed.StartsWith("+");
            string digits = DigitsOnly(trimmed);
            if (string.IsNullOrEmpty(digits)) return string.Empty;
            return leadingPlus ? "+" + digits : digits;
        }

    }
    
}