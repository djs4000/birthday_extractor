// Namespace for the entire Birthday Extractor application.
namespace BirthdayExtractor
{
    // Imports for handling various functionalities like collections, date/time, file I/O, LINQ, text manipulation, and threading.
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.IO;
    using System.Linq;
    using System.Text;
    using System.Threading;
    // Third-party libraries for CSV parsing, phone number handling, and Excel file creation.
    using CsvHelper;
    using CsvHelper.Configuration;
    using PhoneNumbers;
    using ClosedXML.Excel;

    /// <summary>
    /// Strongly-typed container for user input options that drive a single extraction run.
    /// This class holds all the settings and parameters required for the processing logic.
    /// </summary>
    public sealed class ProcOptions
    {
        /// <summary>
        /// Path to the exported customer CSV that will be parsed.
        /// </summary>
        public string CsvPath { get; set; } = string.Empty;
        /// <summary>
        /// Beginning of the inclusive birthday window that should be reported on.
        /// </summary>
        public DateTime Start { get; set; }
        /// <summary>
        /// End of the inclusive birthday window.
        /// </summary>
        public DateTime End { get; set; }
        /// <summary>
        /// Minimum age (turning) that should be considered a child match.
        /// </summary>
        public int MinAge { get; set; } = 3;
        /// <summary>
        /// Maximum age (turning) that still counts a child for celebrations.
        /// </summary>
        public int MaxAge { get; set; } = 14;

        /// <summary>
        /// When true, a filtered CSV export will be produced.
        /// </summary>
        public bool WriteCsv { get; set; } = false;
        /// <summary>
        /// When true, an XLSX workbook with the filtered data will also be produced.
        /// </summary>
        public bool WriteXlsx { get; set; } = true;
        /// <summary>
        /// Directory where any generated exports should be written.
        /// </summary>
        public string OutDir { get; set; } = string.Empty;
        /// <summary>
        /// Optional progress reporter used to surface UI feedback.
        /// </summary>
        public IProgress<int>? Progress { get; set; }
        /// <summary>
        /// Optional log callback for pushing human-readable status messages.
        /// </summary>
        public Action<string>? Log { get; set; }
        /// <summary>
        /// Cancellation token propagated from the UI to abort work gracefully.
        /// </summary>
        public CancellationToken Cancellation { get; set; }

        /// <summary>
        /// Enables libphonenumber validation/normalization when available.
        /// </summary>
        public bool UseLibPhoneNumber { get; set; } = true;
        /// <summary>
        /// Region code supplied to libphonenumber for ambiguous numbers (e.g., "AE" for United Arab Emirates).
        /// </summary>
        public string DefaultRegion { get; set; } = "AE";
    }

    /// <summary>
    /// Result payload describing what was written back to disk for a single run.
    /// This class summarizes the outcome of the processing.
    /// </summary>
    public sealed class ProcResult
    {
        /// <summary>
        /// Number of child rows that survived all filters.
        /// </summary>
        public int KeptCount { get; set; }
        /// <summary>
        /// Full path to the emitted CSV file (if CSV export was requested).
        /// </summary>
        public string? CsvPath { get; set; }
        /// <summary>
        /// Full path to the emitted XLSX file (if XLSX export was requested).
        /// </summary>
        public string? XlsxPath { get; set; }
        /// <summary>
        /// Extracted child/guardian records that can be reused for ERP uploads.
        /// </summary>
        public IReadOnlyList<ExtractedLead> Leads { get; set; } = Array.Empty<ExtractedLead>();
    }

    /// <summary>
    /// Public representation of an extracted child/guardian record used for downstream uploads.
    /// This is a flattened, clean data structure for external systems.
    /// </summary>
    public sealed class ExtractedLead
    {
        public string? ChildFirstName { get; set; }
        public string? ChildLastName { get; set; }
        public string? Email { get; set; }
        public string? Mobile { get; set; }
        public string? NormalizedMobile { get; set; }
        public string? DateOfBirth { get; set; }
        public string? VisitorType { get; set; }
        public string? ParentName { get; set; }
        public string? ParentFirstName { get; set; }
        public string? ParentLastName { get; set; }
        public int Age { get; set; }
        public string BusinessKey { get; set; } = string.Empty;
    }

    /// <summary>
    /// Core engine that parses CSV data, links guardians, and emits exports.
    /// This class is stateless aside from helper buffers, making it reusable.
    /// </summary>
    internal sealed class Processing
    {
        /// <summary>
        /// Lightweight dictionary-backed row representation straight from CsvHelper.
        /// Provides a flexible way to access CSV columns by header name.
        /// </summary>
        private sealed class DynamicRow
        {
            private readonly Dictionary<string, string?> _map;
            public DynamicRow(Dictionary<string, string?> map) => _map = map;
            public string? Get(string key) => _map.TryGetValue(key, out var v) ? v : null;
            public IReadOnlyCollection<string> Keys => _map.Keys;
        }

        /// <summary>
        /// Normalized in-memory record used while correlating guardians and children.
        /// This internal representation standardizes data for easier processing.
        /// </summary>
        private sealed class Row
        {
            public string? FirstName { get; set; }
            public string? LastName { get; set; }
            public string? Email { get; set; }
            public string? Mobile { get; set; }   // original value from CSV
            public string PhoneKey { get; set; } = string.Empty; // normalized phone number (no '+') for matching
            public bool PhoneValid { get; set; }  // validity flag (true if known-valid)
            public string? DobRaw { get; set; }    // original date of birth string
            public DateTime Dob { get; set; }      // parsed date of birth
            public int AgeToday { get; set; }      // calculated age as of today
            public string? VisitorType { get; set; }
        }

        /// <summary>
        /// Data shape used by writers once filtering is complete.
        /// This class holds the final, polished data ready for export.
        /// </summary>
        private sealed class Output
        {
            public string? FirstName { get; set; }
            public string? LastName { get; set; }
            public string? Email { get; set; }
            public string? Mobile { get; set; }
            public string? DateOfBirth { get; set; } // Formatted as yyyy-MM-dd
            public string? VisitorType { get; set; } // Null if column is absent in source
            public string? ParentName { get; set; }
            public string? GuardianFirstName { get; set; }
            public string? GuardianLastName { get; set; }
            public int Age { get; set; }             // The age the child will be on their next birthday
            public int BirthdayDay { get; set; }
            public int BirthdayMonth { get; set; }
            public DateTime NextBirthday { get; set; } // The actual date of the next birthday, used for sorting
            public string? NormalizedMobile { get; set; }   // E.164 or canonical format
            public string BusinessKey { get; set; } = string.Empty; // Unique key for deduplication
        }

        /// <summary>
        /// Runs the full extraction pipeline using the supplied options.
        /// Performs reading, normalization, filtering, and writing in sequence.
        /// </summary>
        public ProcResult Process(ProcOptions o)
        {
            // Ensure the operation hasn't been cancelled before starting.
            o.Cancellation.ThrowIfCancellationRequested();
            // Log initial parameters for debugging and traceability.
            o.Log?.Invoke($"CSV = {o.CsvPath}");
            o.Log?.Invoke($"Start = {o.Start:yyyy-MM-dd}, End = {o.End:yyyy-MM-dd}");

            // Step 1: Read the raw CSV data.
            var rowsDyn = ReadCsv(o.CsvPath, out bool hasVisitorType);
            o.Log?.Invoke($"Loaded {rowsDyn.Length:n0} rows. Visitor Type present: {hasVisitorType}");
            ReportProgress(o, 5);

            // Step 2: Normalize the raw data into a structured 'Row' format.
            var rows = rowsDyn.Select(r => new Row
            {
                FirstName = r.Get("First Name"),
                LastName = r.Get("Last Name"),
                Email = r.Get("Email"),
                Mobile = r.Get("Mobile Number"),
                DobRaw = r.Get("Date of Birth"),
                VisitorType = hasVisitorType ? r.Get("Visitor Type") : null
            }).ToList();

            // Step 3: Normalize phone numbers for reliable matching.
            foreach (var r in rows)
            {
                var (norm, valid) = NormalizePhoneForMatching(r.Mobile, o.UseLibPhoneNumber, o.DefaultRegion);
                r.PhoneKey = norm;   // Digits only, includes country code if known.
                r.PhoneValid = valid;  // True if validated by UAE heuristic or libphonenumber.
            }

            // Step 4: Build maps of adults (>=18 years old) to link children to guardians.
            var adultsByEmail = new Dictionary<string, List<Row>>(StringComparer.OrdinalIgnoreCase);
            var adultsByPhone = new Dictionary<string, List<Row>>(StringComparer.OrdinalIgnoreCase);
            int parsed = 0;
            foreach (var r in rows)
            {
                o.Cancellation.ThrowIfCancellationRequested();
                // Parse date of birth and calculate current age.
                if (TryParseDob(r.DobRaw, out var dob))
                {
                    r.Dob = dob;
                    r.AgeToday = AgeOnDate(dob, DateTime.Today);
                    // If the person is an adult, add them to the lookup maps.
                    if (r.AgeToday >= 18)
                    {
                        if (!string.IsNullOrWhiteSpace(r.Email))
                        {
                            if (!adultsByEmail.TryGetValue(r.Email, out var list)) adultsByEmail[r.Email] = list = new();
                            list.Add(r);
                        }
                        if (!string.IsNullOrEmpty(r.PhoneKey))
                        {
                            // TODO: Consider an option to only include validated phone numbers in the adult map.
                            // This could improve guardian matching accuracy.
                            if (!adultsByPhone.TryGetValue(r.PhoneKey, out var list)) adultsByPhone[r.PhoneKey] = list = new();
                            list.Add(r);
                        }
                    }
                }
                parsed++;
                if (parsed % 5000 == 0) o.Log?.Invoke($"Parsed {parsed:n0}/{rows.Count:n0}...");
            }
            ReportProgress(o, 20);

            // Step 5: Filter for children with birthdays in the specified window and find their guardians.
            var kept = new List<Output>();
            int processed = 0;
            foreach (var r in rows)
            {
                o.Cancellation.ThrowIfCancellationRequested();
                // Ensure we have a valid date of birth for the row.
                DateTime dob;
                if (r.Dob != default) dob = r.Dob;
                else if (TryParseDob(r.DobRaw, out var parsedDob)) dob = parsedDob;
                else continue; // Skip if DOB is invalid.

                // Check if the child's next birthday falls within the desired date window.
                var nextBday = NextBirthdayInWindow(dob, o.Start, o.End);
                if (nextBday == null) continue;

                // Check if the child's age on their next birthday is within the specified age range.
                int ageTurning = nextBday.Value.Year - dob.Year;
                if (ageTurning < o.MinAge || ageTurning > o.MaxAge) continue;

                // --- Find a guardian for the child ---
                Row? guardian = null;
                // Try to find a guardian by matching email first.
                if (!string.IsNullOrWhiteSpace(r.Email) &&
                    adultsByEmail.TryGetValue(r.Email, out var adultsE))
                {
                    guardian = ChooseGuardian(adultsE, preferResident: hasVisitorType);
                }
                // If no guardian found by email, try matching by phone number.
                if (guardian == null &&
                    !string.IsNullOrEmpty(r.PhoneKey) &&
                    adultsByPhone.TryGetValue(r.PhoneKey, out var adultsP))
                {
                    guardian = ChooseGuardian(adultsP, preferResident: hasVisitorType);
                }

                // Format the parent's name.
                string parentName = "";
                if (guardian != null)
                {
                    parentName = string.Join(' ',
                        new[] { guardian.FirstName, guardian.LastName }
                        .Where(s => !string.IsNullOrWhiteSpace(s)));
                }

                // If phone number validation is enabled, skip rows with invalid or empty phone numbers.
                if (o.UseLibPhoneNumber && (string.IsNullOrEmpty(r.PhoneKey) || !r.PhoneValid))
                    continue;

                // --- Create a unique business key for deduplication ---
                // This key helps identify unique individuals even with minor data variations.
                string businessKey;
                if (!string.IsNullOrWhiteSpace(r.PhoneKey))
                {
                    businessKey = $"M:{r.PhoneKey}|DOB:{dob:yyyy-MM-dd}";
                }
                else if (!string.IsNullOrWhiteSpace(r.Email))
                {
                    businessKey = $"E:{r.Email.Trim().ToLowerInvariant()}|DOB:{dob:yyyy-MM-dd}";
                }
                else
                {
                    // Fallback to a name-based key if contact info is missing.
                    var childNameKey = string.Join(" ", new[] { r.FirstName, r.LastName }
                        .Where(s => !string.IsNullOrWhiteSpace(s))
                        .Select(s => s!.Trim().ToLowerInvariant()));
                    businessKey = string.IsNullOrEmpty(childNameKey)
                        ? $"DOB:{dob:yyyy-MM-dd}|ROW:{processed + 1}" // Last resort, use row number.
                        : $"C:{childNameKey}|DOB:{dob:yyyy-MM-dd}";
                }

                // Add the filtered and processed record to the output list.
                kept.Add(new Output
                {
                    FirstName = r.FirstName,
                    LastName = r.LastName,
                    Email = r.Email,
                    Mobile = r.Mobile,             // Original value
                    NormalizedMobile = r.PhoneKey, // Normalized value
                    DateOfBirth = dob.ToString("yyyy-MM-dd"),
                    VisitorType = hasVisitorType ? (r.VisitorType ?? string.Empty) : null,
                    ParentName = parentName,
                    GuardianFirstName = guardian?.FirstName,
                    GuardianLastName = guardian?.LastName,
                    Age = ageTurning,
                    BirthdayDay = dob.Day,
                    BirthdayMonth = dob.Month,
                    NextBirthday = nextBday.Value,
                    BusinessKey = businessKey
                });
                processed++;
                if (processed % 5000 == 0) o.Log?.Invoke($"Processed {processed:n0}/{rows.Count:n0}...");
            }

            // Sort the final list of children by their upcoming birthday.
            kept = kept.OrderBy(k => k.NextBirthday).ThenBy(k => k.BirthdayMonth).ThenBy(k => k.BirthdayDay).ToList();
            o.Log?.Invoke($"Kept {kept.Count:n0} rows.");
            ReportProgress(o, 60);

            // Step 6: Write the output files (CSV and/or XLSX).
            var stamp = $"{o.Start:yyyy-MM-dd}_to_{o.End:yyyy-MM-dd}";
            Directory.CreateDirectory(o.OutDir);
            string? csvOut = null, xlsxOut = null;

            // Prepare the data for potential ERP upload.
            var extractedLeads = kept.Select(k => new ExtractedLead
            {
                ChildFirstName = k.FirstName,
                ChildLastName = k.LastName,
                Email = k.Email,
                Mobile = k.Mobile,
                NormalizedMobile = k.NormalizedMobile,
                DateOfBirth = k.DateOfBirth,
                VisitorType = k.VisitorType,
                ParentName = k.ParentName,
                ParentFirstName = k.GuardianFirstName,
                ParentLastName = k.GuardianLastName,
                Age = k.Age,
                BusinessKey = k.BusinessKey
            }).ToList();

            // Write CSV if requested.
            if (o.WriteCsv)
            {
                var desired = Path.Combine(o.OutDir, $"birthdays_{stamp}.csv");
                csvOut = WriteCsv(kept, desired, hasVisitorType);
                o.Log?.Invoke($"Wrote CSV: {csvOut}");
            }
            ReportProgress(o, 75);

            // Write XLSX if requested.
            if (o.WriteXlsx)
            {
                var desired = Path.Combine(o.OutDir, $"birthdays_{stamp}.xlsx");
                xlsxOut = WriteXlsx(kept, desired, hasVisitorType);
                o.Log?.Invoke($"Wrote XLSX: {xlsxOut}");
            }
            ReportProgress(o, 100);

            // Return the results of the operation.
            return new ProcResult
            {
                KeptCount = kept.Count,
                CsvPath = csvOut,
                XlsxPath = xlsxOut,
                Leads = extractedLeads
            };
        }

        // ---------- Private Helper Methods ----------

        /// <summary>
        /// Strips any non-digit characters from the provided string.
        /// </summary>
        private static string DigitsOnly(string s)
        {
            // TODO: This can be optimized using a Regex or Span-based approach for performance on very large datasets.
            var sb = new System.Text.StringBuilder(s.Length);
            foreach (var ch in s)
                if (ch >= '0' && ch <= '9') sb.Append(ch);
            return sb.ToString();
        }

        /// <summary>
        /// Attempts to convert a UAE mobile number into strict E.164 format (+971...).
        /// Returns true if the number is confidently identified as a UAE mobile number.
        /// Handles various formats like 05..., 9715..., +9715..., etc.
        /// </summary>
        private static bool TryNormalizeUaeMobile(string? input, out string normalized)
        {
            normalized = string.Empty;
            if (string.IsNullOrWhiteSpace(input)) return false;

            // Pre-process the string to remove common formatting characters.
            string compact = input.Trim()
                                .Replace(" ", "")
                                .Replace("-", "")
                                .Replace("(", "")
                                .Replace(")", "");
            string digits = DigitsOnly(compact);

            // Handle international direct dialing prefix "00".
            if (digits.StartsWith("00")) digits = digits.Substring(2);

            // Case 1: Full international format with country code (e.g., 9715########).
            if (digits.Length == 12 && digits.StartsWith("971") && digits[3] == '5')
            {
                normalized = "+971" + digits.Substring(3);
                return true;
            }

            // Case 2: Country code followed by 9-digit mobile number.
            if (digits.StartsWith("971"))
            {
                string tail = digits.Substring(3);
                if (tail.Length == 9 && tail[0] == '5' && IsUaeMobilePrefix(tail.Substring(0, 2)))
                {
                    normalized = "+971" + tail;
                    return true;
                }
            }

            // Case 3: Local 10-digit format (e.g., 05########).
            if (digits.Length == 10 && digits[0] == '0' && digits[1] == '5' && IsUaeMobilePrefix(digits.Substring(1, 2)))
            {
                normalized = "+971" + digits.Substring(1);  // Drop leading 0.
                return true;
            }

            // Case 4: Local 9-digit format (e.g., 5########).
            if (digits.Length == 9 && digits[0] == '5' && IsUaeMobilePrefix(digits.Substring(0, 2)))
            {
                normalized = "+971" + digits;
                return true;
            }

            return false;
        }

        /// <summary>
        /// Checks whether the supplied two-digit prefix belongs to known UAE mobile operator ranges.
        /// </summary>
        private static bool IsUaeMobilePrefix(string twoDigits)
            => twoDigits is "50" or "52" or "53" or "54" or "55" or "56" or "57" or "58";

        /// <summary>
        /// Produces a normalized phone key for deduping guardians and reports its validity.
        /// It first tries a UAE-specific heuristic, then falls back to Google's libphonenumber,
        /// and finally defaults to just the digits if parsing fails.
        /// </summary>
        private static (string normalized, bool valid) NormalizePhoneForMatching(string? input, bool useLib, string defaultRegion)
        {
            if (string.IsNullOrWhiteSpace(input)) return (string.Empty, false);

            // First, try the fast and specific UAE heuristic.
            if (TryNormalizeUaeMobile(input!, out var uaeE164))
            {
                // Strip '+' for the key to ensure consistency.
                return (uaeE164.StartsWith("+") ? uaeE164.Substring(1) : uaeE164, true);
            }

            // If the heuristic fails, use the more general-purpose libphonenumber library if enabled.
            if (useLib)
            {
                var cleanInput = input!.Trim();
                try
                {
                    //var cleanInput = input.Trim();
                    // Libphonenumber expects a '+' for international numbers.
                    if (cleanInput.StartsWith("00"))
                        cleanInput = "+" + cleanInput.Substring(2);

                    var phoneUtil = PhoneNumberUtil.GetInstance();
                    var number = phoneUtil.Parse(cleanInput, defaultRegion);
                    bool valid = phoneUtil.IsValidNumber(number);

                    // Format to E.164 standard and strip '+' for the key.
                    var e164 = phoneUtil.Format(number, PhoneNumberFormat.E164);
                    string noPlus = e164.StartsWith("+") ? e164.Substring(1) : e164;
                    return (noPlus, valid);
                }
                catch (NumberParseException npex)
                {
                    LogRouter.LogException(npex, "Phone parsing failed: " + cleanInput);
                }
            }

            // Fallback: if all else fails, just use the digits from the input.
            string trimmed = input!.Trim();
            string digits = DigitsOnly(trimmed);
            if (digits.Length == 0) return (string.Empty, false);

            // We cannot claim validity for this fallback.
            return (digits, false);
        }

        /// <summary>
        /// Detects whether the target file is locked by another process before writing.
        /// </summary>
        private static bool IsFileLocked(string path)
        {
            try
            {
                if (!File.Exists(path)) return false;
                // Attempt to open the file with exclusive access.
                using var _ = File.Open(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
                return false;
            }
            catch (IOException ioex)
            {
                LogRouter.LogException(ioex, "File lock detection failed");
                return true; // File is locked.
            }
            catch (Exception ex)
            {
                LogRouter.LogException(ex, "Unexpected file lock error");
                return true; // Another unexpected error occurred.
            }
        }

        /// <summary>
        /// Generates a conflict-free output path by appending numbers or a timestamp if the desired path is taken.
        /// Example: "file.xlsx" -> "file (1).xlsx" -> "file_20250927_123000.xlsx"
        /// </summary>
        private static string GetSafeOutputPath(string desiredPath)
        {
            // If the path is not locked and doesn't exist, we can use it directly.
            if (!File.Exists(desiredPath) && !IsFileLocked(desiredPath))
                return desiredPath;

            // If conflicted, try numbered suffixes like "filename (1).ext".
            var dir = Path.GetDirectoryName(desiredPath) ?? ".";
            var name = Path.GetFileNameWithoutExtension(desiredPath);
            var ext = Path.GetExtension(desiredPath);
            for (int i = 1; i <= 99; i++)
            {
                var candidate = Path.Combine(dir, $"{name} ({i}){ext}");
                if (!File.Exists(candidate) && !IsFileLocked(candidate))
                    return candidate;
            }

            // As a last resort, append a timestamp to ensure uniqueness.
            var stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            var tsPath = Path.Combine(dir, $"{name}_{stamp}{ext}");
            return tsPath;
        }

        /// <summary>
        /// Reads the raw CSV into dynamic rows, automatically detecting the delimiter and encoding.
        /// It also checks for the presence of the optional "Visitor Type" column.
        /// </summary>
        private static DynamicRow[] ReadCsv(string path, out bool hasVisitorType)
        {
            using var fs = File.OpenRead(path);
            // Detect encoding (e.g., UTF-8 with BOM) or default to UTF-8.
            using var sr = new StreamReader(fs, DetectEncoding(fs) ?? new UTF8Encoding(false), detectEncodingFromByteOrderMarks: true);

            // Auto-detect delimiter by sampling the first line.
            string? sample = sr.ReadLine();
            if (sample == null) throw new InvalidOperationException("Cannot process an empty or header-only CSV file.");
            char delimiter = sample.Contains(';') && !sample.Contains(',') ? ';' : ',';
            sr.DiscardBufferedData(); fs.Seek(0, SeekOrigin.Begin);

            // Configure CsvHelper for robust parsing.
            var config = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                Delimiter = delimiter.ToString(),
                IgnoreBlankLines = true,
                DetectColumnCountChanges = false, // Be lenient with column counts.
                BadDataFound = null, // Ignore bad data fields.
                TrimOptions = TrimOptions.Trim,
                MissingFieldFound = null // Handle missing fields gracefully.
            };

            var list = new List<DynamicRow>();
            using var csv = new CsvReader(sr, config);
            csv.Read();
            csv.ReadHeader();
            var headers = (csv.HeaderRecord ?? Array.Empty<string>()).ToList();
            hasVisitorType = headers.Any(h => string.Equals(h, "Visitor Type", StringComparison.OrdinalIgnoreCase));

            // Read each row into a dictionary-based DynamicRow.
            while (csv.Read())
            {
                var dict = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);
                foreach (var h in headers)
                {
                    string? v = null;
                    try
                    {
                        v = csv.GetField(h);
                    }
                    catch (Exception ex)
                    {
                        LogRouter.LogException(ex, $"Failed to read CSV field '{h}'");
                    }
                    dict[h] = string.IsNullOrWhiteSpace(v) ? null : v;
                }
                list.Add(new DynamicRow(dict));
            }
            return list.ToArray();
        }

        /// <summary>
        /// Peeks at the beginning of a stream to detect the file encoding from Byte Order Marks (BOM).
        /// </summary>
        private static Encoding? DetectEncoding(Stream s)
        {
            long pos = s.Position;
            Span<byte> bom = stackalloc byte[4];
            int read = s.Read(bom);
            s.Seek(pos, SeekOrigin.Begin); // Reset stream position.

            if (read >= 3 && bom[0] == 0xEF && bom[1] == 0xBB && bom[2] == 0xBF) return new UTF8Encoding(true); // UTF-8 BOM
            if (read >= 2 && bom[0] == 0xFF && bom[1] == 0xFE) return Encoding.Unicode; // UTF-16 LE
            if (read >= 2 && bom[0] == 0xFE && bom[1] == 0xFF) return Encoding.BigEndianUnicode; // UTF-16 BE
            return null; // No BOM detected.
        }

        /// <summary>
        /// Attempts to parse a Date of Birth from a wide variety of common string formats.
        /// </summary>
        private static bool TryParseDob(string? s, out DateTime dob)
        {
            dob = default;
            if (string.IsNullOrWhiteSpace(s)) return false;

            // A list of exact formats to try first for performance and accuracy.
            string[] fmts = {
                "yyyy-MM-dd", "dd/MM/yyyy", "MM/dd/yyyy", "yyyy/MM/dd", "dd-MM-yyyy", "MM-dd-yyyy",
                "d/M/yyyy", "M/d/yyyy", "dd.MM.yyyy", "d.M.yyyy", "d-MMM-yy", "d MMM yyyy"
            };

            // Try parsing with the exact formats first.
            if (DateTime.TryParseExact(s, fmts, CultureInfo.InvariantCulture, DateTimeStyles.None, out dob)) return true;
            // Fallback to more general parsing with invariant culture.
            if (DateTime.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.None, out dob)) return true;
            // Final fallback with the system's current culture.
            if (DateTime.TryParse(s, CultureInfo.CurrentCulture, DateTimeStyles.None, out dob)) return true;

            return false;
        }

        /// <summary>
        /// Calculates the date of the next birthday for a person that falls within a given date window.
        /// Returns null if the next birthday is outside the window.
        /// </summary>
        private static DateTime? NextBirthdayInWindow(DateTime dob, DateTime start, DateTime end)
        {
            // Helper to map a birthday to a specific year, handling leap years.
            static DateTime MapToYear(DateTime birth, int year)
            {
                int day = birth.Day; int month = birth.Month;
                // Handle leap day birthdays in non-leap years by moving them to Feb 28.
                if (month == 2 && day == 29 && !DateTime.IsLeapYear(year)) day = 28;
                return new DateTime(year, month, day);
            }

            var s = start.Date; var e = end.Date;
            // Candidate birthday in the start year.
            var cand1 = MapToYear(dob, s.Year);

            // Standard case: window is within a single year.
            if (e >= s)
            {
                if (cand1 >= s && cand1 <= e) return cand1;
                return null;
            }

            // Edge case: window crosses a year boundary (e.g., Dec 15 to Jan 15).
            if (cand1 >= s) return cand1; // Birthday is in the later part of the start year.
            var cand2 = MapToYear(dob, s.Year + 1); // Check the next year.
            if (cand2 <= e.AddYears(1)) return cand2; // Birthday is in the early part of the end year.

            return null;
        }

        /// <summary>
        /// Calculates a person's age on a specified reference date.
        /// </summary>
        private static int AgeOnDate(DateTime dob, DateTime refDate)
        {
            int age = refDate.Year - dob.Year;
            // Decrement age if the birthday has not yet occurred in the reference year.
            if (refDate.Month < dob.Month || (refDate.Month == dob.Month && refDate.Day < dob.Day)) age--;
            return age;
        }

        /// <summary>
        /// From a list of potential adult guardians, chooses the most suitable one.
        /// It can be configured to prefer "Resident" visitor types if available.
        /// </summary>
        private static Row? ChooseGuardian(List<Row> adults, bool preferResident)
        {
            IEnumerable<Row> pool = adults;
            // If we prefer residents and have that data, filter the pool.
            if (preferResident)
            {
                var residents = adults.Where(a => string.Equals(a.VisitorType, "Resident", StringComparison.OrdinalIgnoreCase)).ToList();
                if (residents.Count > 0) pool = residents;
            }
            // Prefer adults with a last name, as they are more likely to be fully registered.
            return pool.OrderByDescending(a => !string.IsNullOrWhiteSpace(a.LastName)).FirstOrDefault();
        }

        /// <summary>
        /// Writes the filtered results to a CSV file, adding headers dynamically based on source columns.
        /// </summary>
        private static string WriteCsv(List<Output> rows, string path, bool hasVisitorType)
        {
            var savePath = GetSafeOutputPath(path);
            using var fs = File.Create(savePath);
            using var sw = new StreamWriter(fs, new UTF8Encoding(false)); // UTF-8 without BOM
            using var csv = new CsvWriter(sw, CultureInfo.InvariantCulture);

            // Define headers, including the optional "Visitor Type".
            var headers = new List<string> { "First Name", "Last Name", "Email", "Mobile Number", "Mobile (Normalized)", "Date of Birth" };
            if (hasVisitorType) headers.Add("Visitor Type");
            headers.AddRange(new[] { "Parent Name", "Age", "Birthday day", "Birthday month" });

            // Write header row.
            foreach (var h in headers) csv.WriteField(h);
            csv.NextRecord();

            // Write data rows.
            foreach (var r in rows)
            {
                csv.WriteField(r.FirstName);
                csv.WriteField(r.LastName);
                csv.WriteField(r.Email);
                csv.WriteField(r.Mobile);
                csv.WriteField(r.NormalizedMobile);
                csv.WriteField(r.DateOfBirth);
                if (hasVisitorType) csv.WriteField(r.VisitorType);
                csv.WriteField(r.ParentName);
                csv.WriteField(r.Age);
                csv.WriteField(r.BirthdayDay);
                csv.WriteField(r.BirthdayMonth);
                csv.NextRecord();
            }
            return savePath;
        }

        /// <summary>
        /// Writes the filtered results to an XLSX workbook with professional table formatting.
        /// </summary>
        private static string WriteXlsx(List<Output> rows, string path, bool hasVisitorType)
        {
            var savePath = GetSafeOutputPath(path);
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Birthdays");

            // Define headers, including the optional "Visitor Type".
            var headers = new List<string> { "First Name", "Last Name", "Email", "Mobile Number", "Mobile (Normalized)", "Date of Birth" };
            if (hasVisitorType) headers.Add("Visitor Type");
            headers.AddRange(new[] { "Parent Name", "Age", "Birthday day", "Birthday month" });

            // Write header row.
            for (int c = 0; c < headers.Count; c++) ws.Cell(1, c + 1).Value = headers[c];

            // Write data rows.
            for (int r = 0; r < rows.Count; r++)
            {
                var row = rows[r]; int c = 1;
                ws.Cell(r + 2, c++).Value = row.FirstName ?? string.Empty;
                ws.Cell(r + 2, c++).Value = row.LastName ?? string.Empty;
                ws.Cell(r + 2, c++).Value = row.Email ?? string.Empty;
                ws.Cell(r + 2, c++).Value = row.Mobile ?? string.Empty;
                ws.Cell(r + 2, c++).Value = row.NormalizedMobile ?? string.Empty;
                ws.Cell(r + 2, c++).Value = row.DateOfBirth ?? string.Empty;
                if (hasVisitorType) ws.Cell(r + 2, c++).Value = row.VisitorType ?? string.Empty;
                ws.Cell(r + 2, c++).Value = row.ParentName ?? string.Empty;
                ws.Cell(r + 2, c++).Value = row.Age;
                ws.Cell(r + 2, c++).Value = row.BirthdayDay;
                ws.Cell(r + 2, c++).Value = row.BirthdayMonth;
            }

            // Apply table formatting for better readability.
            var lastRow = Math.Max(1, rows.Count + 1);
            var lastCol = headers.Count;
            var range = ws.Range(1, 1, lastRow, lastCol);
            var table = range.CreateTable("BirthdaysTable");
            table.Theme = XLTableTheme.TableStyleMedium2;

            // Apply number formatting to numeric columns.
            if (rows.Count > 0)
            {
                var ageCol = headers.IndexOf("Age") + 1;
                var dayCol = headers.IndexOf("Birthday day") + 1;
                var monthCol = headers.IndexOf("Birthday month") + 1;
                if (ageCol > 0) ws.Column(ageCol).Style.NumberFormat.Format = "0";
                if (dayCol > 0) ws.Column(dayCol).Style.NumberFormat.Format = "0";
                if (monthCol > 0) ws.Column(monthCol).Style.NumberFormat.Format = "0";
            }

            // Auto-fit column widths to content.
            ws.Columns().AdjustToContents();
            wb.SaveAs(savePath);
            return savePath;
        }

        /// <summary>
        /// Helper for safely reporting progress increments back to the UI thread.
        /// </summary>
        private static void ReportProgress(ProcOptions o, int pct) => o.Progress?.Report(pct);
    }
}