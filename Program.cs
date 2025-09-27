using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Windows.Forms;

namespace BirthdayExtractor
{
    /// <summary>
    /// Entry point for the Birthday Extractor Windows Forms application.
    /// Responsible for bootstrapping WinForms defaults and the main UI.
    /// </summary>
    internal static class Program
    {
        /// <summary>
        /// Boots the UI thread, applies visual/high DPI settings, and opens <see cref="MainForm"/>.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            if (TryRunSilent(args))
            {
                return;
            }

            Application.EnableVisualStyles();
            Application.SetHighDpiMode(HighDpiMode.SystemAware);
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }

        private static bool TryRunSilent(string[] args)
        {
            if (args is null || args.Length == 0)
            {
                return false;
            }

            var parsed = ParseArguments(args);
            if (parsed.Count == 0)
            {
                return false;
            }

            if (parsed.ContainsKey("help") || parsed.ContainsKey("?"))
            {
                PrintUsage();
                Environment.ExitCode = 0;
                return true;
            }

            var silentRequested = GetFlag(parsed, "silent") || GetFlag(parsed, "cli") || GetFlag(parsed, "headless");
            if (!silentRequested)
            {
                return false;
            }

            if (!parsed.TryGetValue("csv", out var csvPath) || string.IsNullOrWhiteSpace(csvPath))
            {
                Console.Error.WriteLine("ERROR: --csv <path> is required when running in silent mode.");
                PrintUsage();
                Environment.ExitCode = 1;
                return true;
            }

            if (!File.Exists(csvPath))
            {
                Console.Error.WriteLine($"ERROR: CSV file not found: {csvPath}");
                Environment.ExitCode = 1;
                return true;
            }

            if (!parsed.TryGetValue("start", out var startText) || !TryParseDate(startText, out var start))
            {
                Console.Error.WriteLine("ERROR: --start <yyyy-MM-dd> is required when running in silent mode.");
                Environment.ExitCode = 1;
                return true;
            }

            if (!parsed.TryGetValue("end", out var endText) || !TryParseDate(endText, out var end))
            {
                Console.Error.WriteLine("ERROR: --end <yyyy-MM-dd> is required when running in silent mode.");
                Environment.ExitCode = 1;
                return true;
            }

            if (end < start)
            {
                Console.Error.WriteLine("ERROR: --end must be on or after --start.");
                Environment.ExitCode = 1;
                return true;
            }

            var outputDir = parsed.TryGetValue("out", out var outDir) && !string.IsNullOrWhiteSpace(outDir)
                ? outDir!
                : (Path.GetDirectoryName(Path.GetFullPath(csvPath)) ?? Environment.CurrentDirectory);

            AppConfig config;
            try
            {
                config = ConfigStore.LoadOrCreate() ?? new AppConfig();
            }
            catch
            {
                config = new AppConfig();
            }

            var writeCsv = GetFlag(parsed, "csv-out", config.DefaultWriteCsv) && !GetFlag(parsed, "no-csv-out");
            var writeXlsx = GetFlag(parsed, "xlsx-out", config.DefaultWriteXlsx) && !GetFlag(parsed, "no-xlsx-out");
            var quiet = GetFlag(parsed, "quiet");

            if (!parsed.TryGetValue("min-age", out var minAgeText) || !int.TryParse(minAgeText, out var minAge))
            {
                minAge = config.MinAge;
            }

            if (!parsed.TryGetValue("max-age", out var maxAgeText) || !int.TryParse(maxAgeText, out var maxAge))
            {
                maxAge = config.MaxAge;
            }

            var defaultRegion = parsed.TryGetValue("default-region", out var region) && !string.IsNullOrWhiteSpace(region)
                ? region!
                : config.DefaultRegion;

            var useLibPhoneNumber = GetFlag(parsed, "libphonenumber", config.UseLibPhoneNumber) && !GetFlag(parsed, "no-libphonenumber");

            try
            {
                Directory.CreateDirectory(outputDir);
                var processor = new Processing();

                var result = processor.Process(new ProcOptions
                {
                    CsvPath = csvPath,
                    Start = start,
                    End = end,
                    OutDir = outputDir,
                    WriteCsv = writeCsv,
                    WriteXlsx = writeXlsx,
                    MinAge = minAge,
                    MaxAge = maxAge,
                    UseLibPhoneNumber = useLibPhoneNumber,
                    DefaultRegion = defaultRegion,
                    Cancellation = default,
                    Log = quiet ? null : (Action<string>)(m => Console.WriteLine($"{DateTime.Now:HH:mm:ss}  {m}"))
                });

                if (!quiet)
                {
                    Console.WriteLine($"Completed. Kept {result.KeptCount} rows.");
                    if (result.CsvPath is not null) Console.WriteLine($"CSV : {result.CsvPath}");
                    if (result.XlsxPath is not null) Console.WriteLine($"XLSX: {result.XlsxPath}");
                }

                TryLogHistory(config, csvPath, start, end, result);
                Environment.ExitCode = 0;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("ERROR: " + ex.Message);
                Environment.ExitCode = 1;
            }

            return true;
        }

        private static Dictionary<string, string?> ParseArguments(string[] args)
        {
            var result = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < args.Length; i++)
            {
                var arg = args[i];
                if (string.IsNullOrWhiteSpace(arg))
                {
                    continue;
                }

                if (!IsSwitch(arg))
                {
                    continue;
                }

                var trimmed = arg.TrimStart('-', '/');
                string? value = null;

                var eqIdx = trimmed.IndexOf('=');
                if (eqIdx >= 0)
                {
                    value = trimmed[(eqIdx + 1)..];
                    trimmed = trimmed[..eqIdx];
                }
                else if (i + 1 < args.Length && !IsSwitch(args[i + 1]))
                {
                    value = args[++i];
                }
                else
                {
                    value = "true";
                }

                if (!string.IsNullOrWhiteSpace(trimmed))
                {
                    result[trimmed] = value;
                }
            }

            return result;
        }

        private static bool IsSwitch(string arg)
            => arg.StartsWith("--", StringComparison.Ordinal) || arg.StartsWith("-", StringComparison.Ordinal) || arg.StartsWith("/", StringComparison.Ordinal);

        private static bool GetFlag(IDictionary<string, string?> values, string key, bool defaultValue = false)
        {
            if (!values.TryGetValue(key, out var raw) || string.IsNullOrWhiteSpace(raw))
            {
                return defaultValue;
            }

            if (bool.TryParse(raw, out var parsedBool))
            {
                return parsedBool;
            }

            return raw switch
            {
                "1" => true,
                "0" => false,
                "yes" or "y" => true,
                "no" or "n" => false,
                _ => defaultValue
            };
        }

        private static bool TryParseDate(string? text, out DateTime value)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                value = default;
                return false;
            }

            return DateTime.TryParseExact(
                text.Trim(),
                new[] { "yyyy-MM-dd", "yyyy/MM/dd" },
                CultureInfo.InvariantCulture,
                DateTimeStyles.AssumeLocal | DateTimeStyles.AllowWhiteSpaces,
                out value);
        }

        private static void TryLogHistory(AppConfig cfg, string csvPath, DateTime start, DateTime end, ProcResult result)
        {
            try
            {
                var csvName = Path.GetFileName(csvPath);
                var sha = ConfigStore.ComputeSha256(csvPath);
                cfg.History.Add(new ProcessedWindow
                {
                    Start = start,
                    End = end,
                    CsvName = csvName,
                    CsvSha256 = sha,
                    RowCount = result.KeptCount,
                    ProcessedAt = DateTime.Now
                });
                ConfigStore.Save(cfg);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("WARN: Failed to record run history: " + ex.Message);
            }
        }

        private static void PrintUsage()
        {
            Console.WriteLine("Birthday Extractor silent mode");
            Console.WriteLine("Usage: BirthdayExtractor.exe --silent --csv <path> --start <yyyy-MM-dd> --end <yyyy-MM-dd> [options]");
            Console.WriteLine();
            Console.WriteLine("Options:");
            Console.WriteLine("  --out <dir>             Output directory (defaults to CSV folder)");
            Console.WriteLine("  --csv-out               Force CSV export (overrides config)");
            Console.WriteLine("  --no-csv-out            Disable CSV export");
            Console.WriteLine("  --xlsx-out              Force XLSX export (overrides config)");
            Console.WriteLine("  --no-xlsx-out           Disable XLSX export");
            Console.WriteLine("  --min-age <n>           Minimum age filter (defaults to config)");
            Console.WriteLine("  --max-age <n>           Maximum age filter (defaults to config)");
            Console.WriteLine("  --libphonenumber        Enable libphonenumber normalization");
            Console.WriteLine("  --no-libphonenumber     Disable libphonenumber normalization");
            Console.WriteLine("  --default-region <code> Override libphonenumber region");
            Console.WriteLine("  --quiet                 Suppress console logging");
            Console.WriteLine("  --help                  Show this help text");
        }
    }
}
