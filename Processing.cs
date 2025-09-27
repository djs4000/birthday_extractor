using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using CsvHelper;
using CsvHelper.Configuration;
using PhoneNumbers;
using ClosedXML.Excel;
namespace BirthdayExtractor
{
    /// <summary>
    /// Strongly-typed container for user input options that drive a single extraction run.
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
        /// Region code supplied to libphonenumber for ambiguous numbers.
        /// </summary>

        public string DefaultRegion { get; set; } = "AE";
    }
    /// <summary>
    /// Result payload describing what was written back to disk for a single run.
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
    /// Stateless aside from helper buffers.
    /// </summary>
    internal sealed class Processing
    {
        /// <summary>
        /// Lightweight dictionary-backed row representation straight from CsvHelper.
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
        /// </summary>
        private sealed class Row
        {
            public string? FirstName { get; set; }
            public string? LastName  { get; set; }
            public string? Email     { get; set; }
            public string? Mobile    { get; set; }   // original
            public string  PhoneKey  { get; set; } = string.Empty; // normalized (no '+')
            public bool    PhoneValid { get; set; }  // validity (true if known-valid)
            public string? DobRaw    { get; set; }
            public DateTime Dob      { get; set; }
            public int AgeToday      { get; set; }
            public string? VisitorType { get; set; }
        }
        /// <summary>
        /// Data shape used by writers once filtering is complete.
        /// </summary>
        private sealed class Output
        {
            public string? FirstName { get; set; }
            public string? LastName { get; set; }
            public string? Email { get; set; }
            public string? Mobile { get; set; }
            public string? DateOfBirth { get; set; } // yyyy-MM-dd
            public string? VisitorType { get; set; } // null if column absent
            public string? ParentName { get; set; }
            public string? GuardianFirstName { get; set; }
            public string? GuardianLastName { get; set; }
            public int Age { get; set; }             // turning age
            public int BirthdayDay { get; set; }
            public int BirthdayMonth { get; set; }
            public DateTime NextBirthday { get; set; } // for sorting only
            public string? NormalizedMobile { get; set; }   // NEW
            public string BusinessKey { get; set; } = string.Empty;
        }
        /// <summary>
        /// Runs the full extraction pipeline using the supplied options.
        /// Performs reading, normalization, filtering, and writing in sequence.
        /// </summary>
        public ProcResult Process(ProcOptions o)
        {

            o.Cancellation.ThrowIfCancellationRequested();
            o.Log?.Invoke($"CSV = {o.CsvPath}");
            o.Log?.Invoke($"Start = {o.Start:yyyy-MM-dd}, End = {o.End:yyyy-MM-dd}");
            var rowsDyn = ReadCsv(o.CsvPath, out bool hasVisitorType);
            o.Log?.Invoke($"Loaded {rowsDyn.Length:n0} rows. Visitor Type present: {hasVisitorType}");
            ReportProgress(o, 5);
            // Normalize
            var rows = rowsDyn.Select(r => new Row
            {
                FirstName = r.Get("First Name"),
                LastName = r.Get("Last Name"),
                Email = r.Get("Email"),
                Mobile = r.Get("Mobile Number"),
                DobRaw = r.Get("Date of Birth"),
                VisitorType = hasVisitorType ? r.Get("Visitor Type") : null
            }).ToList();
            // Normalize phone keys for matching
            foreach (var r in rows)
            {
                var (norm, valid) = NormalizePhoneForMatching(r.Mobile, o.UseLibPhoneNumber, o.DefaultRegion);
                r.PhoneKey   = norm;   // digits only, includes country code if known
                r.PhoneValid = valid;  // true if UAE-heuristic or libphonenumber validated it
            }
            // Build adult maps (>=18) so younger guests can inherit guardian contact info
            var adultsByEmail = new Dictionary<string, List<Row>>(StringComparer.OrdinalIgnoreCase);
            var adultsByPhone = new Dictionary<string, List<Row>>(StringComparer.OrdinalIgnoreCase);
            int parsed = 0;
            foreach (var r in rows)
            {
                o.Cancellation.ThrowIfCancellationRequested();
                if (TryParseDob(r.DobRaw, out var dob))
                {
                    r.Dob = dob;
                    r.AgeToday = AgeOnDate(dob, DateTime.Today);
                    if (r.AgeToday >= 18)
                    {
                        if (!string.IsNullOrWhiteSpace(r.Email))
                        {
                            if (!adultsByEmail.TryGetValue(r.Email, out var list)) adultsByEmail[r.Email] = list = new();
                            list.Add(r);
                        }
                        if (!string.IsNullOrEmpty(r.PhoneKey))
                        {
                            // If you want to exclude invalid phones from adult map when lib is on, uncomment next line:
                            // If libphonenumber is on, only keep rows with a valid normalized phone
                            if (!adultsByPhone.TryGetValue(r.PhoneKey, out var list)) adultsByPhone[r.PhoneKey] = list = new();
                            list.Add(r);
                            // skipPhone:;
                        }
                    }
                }
                parsed++;
                if (parsed % 5000 == 0) o.Log?.Invoke($"Parsed {parsed:n0}/{rows.Count:n0}...");
            }
            ReportProgress(o, 20);
            // Filter children and compute output
            var kept = new List<Output>();
            int processed = 0;
            foreach (var r in rows)
            {
                o.Cancellation.ThrowIfCancellationRequested();
                // Resolve DOB
                DateTime dob;
                if (r.Dob != default) dob = r.Dob;
                else if (TryParseDob(r.DobRaw, out var parsedDob)) dob = parsedDob;
                else continue;
                // In-window birthday + turning age
                var nextBday = NextBirthdayInWindow(dob, o.Start, o.End);
                if (nextBday == null) continue;
                int ageTurning = nextBday.Value.Year - dob.Year; // age they will turn
                if (ageTurning < o.MinAge || ageTurning > o.MaxAge) continue;
                // --- choose guardian (email preferred, then normalized phone) ---
                Row? guardian = null;
                if (!string.IsNullOrWhiteSpace(r.Email) &&
                    adultsByEmail.TryGetValue(r.Email, out var adultsE))
                {
                    guardian = ChooseGuardian(adultsE, preferResident: hasVisitorType);
                }
                if (guardian == null &&
                    !string.IsNullOrEmpty(r.PhoneKey) &&
                    adultsByPhone.TryGetValue(r.PhoneKey, out var adultsP))
                {
                    guardian = ChooseGuardian(adultsP, preferResident: hasVisitorType);
                }
                // --- build parentName string ---
                string parentName = "";
                if (guardian != null)
                {
                    parentName = string.Join(' ',
                        new[] { guardian.FirstName, guardian.LastName }
                        .Where(s => !string.IsNullOrWhiteSpace(s)));
                }
                if (o.UseLibPhoneNumber && (string.IsNullOrEmpty(r.PhoneKey) || !r.PhoneValid))
                    continue;
                // Build dedupe/business key
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
                    var childNameKey = string.Join(" ", new[] { r.FirstName, r.LastName }
                        .Where(s => !string.IsNullOrWhiteSpace(s))
                        .Select(s => s!.Trim().ToLowerInvariant()));
                    businessKey = string.IsNullOrEmpty(childNameKey)
                        ? $"DOB:{dob:yyyy-MM-dd}|ROW:{processed + 1}"
                        : $"C:{childNameKey}|DOB:{dob:yyyy-MM-dd}";
                }
                // Add to output
                kept.Add(new Output
                {
                    FirstName     = r.FirstName,
                    LastName      = r.LastName,
                    Email         = r.Email,
                    Mobile        = r.Mobile,             // original value
                    NormalizedMobile = r.PhoneKey,         // normalized (E.164 for UAE, or canonical fallback)
                    DateOfBirth   = dob.ToString("yyyy-MM-dd"),
                    VisitorType   = hasVisitorType ? (r.VisitorType ?? string.Empty) : null,
                    ParentName    = parentName,          // <-- now defined
                    GuardianFirstName = guardian?.FirstName,
                    GuardianLastName  = guardian?.LastName,
                    Age           = ageTurning,
                    BirthdayDay   = dob.Day,
                    BirthdayMonth = dob.Month,
                    NextBirthday  = nextBday.Value,
                    BusinessKey   = businessKey
                });
                processed++;
                if (processed % 5000 == 0) o.Log?.Invoke($"Processed {processed:n0}/{rows.Count:n0}...");
            }
            kept = kept.OrderBy(k => k.NextBirthday).ThenBy(k => k.BirthdayMonth).ThenBy(k => k.BirthdayDay).ToList();
            o.Log?.Invoke($"Kept {kept.Count:n0} rows.");
            ReportProgress(o, 60);
            // Write outputs
            var stamp = $"{o.Start:yyyy-MM-dd}_to_{o.End:yyyy-MM-dd}";
            Directory.CreateDirectory(o.OutDir);
            string? csvOut = null, xlsxOut = null;
            var extractedLeads = kept.Select(k => new ExtractedLead
            {
                ChildFirstName   = k.FirstName,
                ChildLastName    = k.LastName,
                Email            = k.Email,
                Mobile           = k.Mobile,
                NormalizedMobile = k.NormalizedMobile,
                DateOfBirth      = k.DateOfBirth,
                VisitorType      = k.VisitorType,
                ParentName       = k.ParentName,
                ParentFirstName  = k.GuardianFirstName,
                ParentLastName   = k.GuardianLastName,
                Age              = k.Age,
                BusinessKey      = k.BusinessKey
            }).ToList();
            if (o.WriteCsv)
            {
                var desired = Path.Combine(o.OutDir, $"birthdays_{stamp}.csv");
                csvOut = WriteCsv(kept, desired, hasVisitorType);
                o.Log?.Invoke($"Wrote CSV: {csvOut}");
            }
            ReportProgress(o, 75);
            if (o.WriteXlsx)
            {
                var desired = Path.Combine(o.OutDir, $"birthdays_{stamp}.xlsx");
                xlsxOut = WriteXlsx(kept, desired, hasVisitorType);
                o.Log?.Invoke($"Wrote XLSX: {xlsxOut}");
            }
            ReportProgress(o, 100);
            return new ProcResult
            {
                KeptCount = kept.Count,
                CsvPath = csvOut,
                XlsxPath = xlsxOut,
                Leads = extractedLeads
            };
        }
        // ---------- helpers ----------
        /// <summary>
        /// Strips any non-digit characters from the provided string.
        /// </summary>
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
        /// Attempts to convert a UAE mobile number into strict E.164 form.
        /// Returns true when confident about the country/operator.
        /// </summary>
        private static bool TryNormalizeUaeMobile(string? input, out string normalized)
        {
            normalized = string.Empty;
            if (string.IsNullOrWhiteSpace(input)) return false;
            // strip common punctuation/spacing first (we'll also re-extract digits)
            string compact = input.Trim()
                                .Replace(" ", "")
                                .Replace("-", "")
                                .Replace("(", "")
                                .Replace(")", "");
            string digits = DigitsOnly(compact);
            // Handle IDD "00" prefix
            if (digits.StartsWith("00")) digits = digits.Substring(2);
            // Case 1: 971 + 9 digits and must be a mobile starting with '5'
            // Examples: 9715######## (length 12)
            if (digits.Length == 12 && digits.StartsWith("971") && digits[3] == '5')
            {
                normalized = "+971" + digits.Substring(3); // -> +9715########
                return true;
            }
            // Case 2: "971" followed by exactly 9 digits, first is 5
            if (digits.StartsWith("971"))
            {
                string tail = digits.Substring(3);
                if (tail.Length == 9 && tail[0] == '5' && IsUaeMobilePrefix(tail.Substring(0, 2)))
                {
                    normalized = "+971" + tail;
                    return true;
                }
            }
            // Case 3: local 10 digits "05########"
            if (digits.Length == 10 && digits[0] == '0' && digits[1] == '5' && IsUaeMobilePrefix(digits.Substring(1, 2)))
            {
                normalized = "+971" + digits.Substring(1);  // drop leading 0
                return true;
            }
            // Case 4: local 9 digits "5########"
            if (digits.Length == 9 && digits[0] == '5' && IsUaeMobilePrefix(digits.Substring(0, 2)))
            {
                normalized = "+971" + digits;
                return true;
            }
            return false;
        }
        /// <summary>
        /// Checks whether the supplied prefix belongs to known UAE mobile ranges.
        /// </summary>
        private static bool IsUaeMobilePrefix(string twoDigits)
            => twoDigits is "50" or "52" or "53" or "54" or "55" or "56" or "57" or "58";
        /// Normalization for matching across all phones:
        /// - If UAE mobile pattern recognized -> E.164 (+9715########)
        /// - Else: return canonical digits (with leading '+' if present) for stable matching, no validity asserted.
        /// </summary>
        /// <summary>
        /// Produces a normalized phone key for deduping guardians and reports validity.
        /// Falls back to digits only when advanced parsing fails.
        /// </summary>
        private static (string normalized, bool valid) NormalizePhoneForMatching(string? input, bool useLib, string defaultRegion)
        {
            if (string.IsNullOrWhiteSpace(input)) return (string.Empty, false);
            // UAE heuristic first
            if (TryNormalizeUaeMobile(input!, out var uaeE164))
            {
                // strip '+' for storage/output
                return (uaeE164.StartsWith("+") ? uaeE164.Substring(1) : uaeE164, true);
            }
            if (useLib)
            {
                try
                {
                    var cleanInput = input.Trim();
                    // Per libphonenumber convention, replace leading "00" with "+" to signify an international number.
                    // If it doesn't start with "+", it will be parsed relative to the defaultRegion.
                    if (cleanInput.StartsWith("00"))
                        cleanInput = "+" + cleanInput.Substring(2);

                    var phoneUtil = PhoneNumberUtil.GetInstance();
                    var number = phoneUtil.Parse(cleanInput, defaultRegion);
                    bool valid = phoneUtil.IsValidNumber(number);

                    // Format as E164 and strip '+' for the key
                    var e164 = phoneUtil.Format(number, PhoneNumberFormat.E164);
                    string noPlus = e164.StartsWith("+") ? e164.Substring(1) : e164;
                    return (noPlus, valid);
                }
                catch (NumberParseException ex)
                {
                    // The log delegate is better for UI/CLI consistency.
                    // log?.Invoke($"[WARN] libphonenumber failed to parse '{input}': {ex.Message}");
                }
            }
            // Fallback: keep digits (and if there was a '+', it's already in the digits)
            string trimmed = input!.Trim();
            string digits = DigitsOnly(trimmed);
            if (digits.Length == 0) return (string.Empty, false);
            // We don’t claim validity here
            return (digits, false);
        }
        /// <summary>
        /// Detects whether the target path is already open by another process.
        /// </summary>
        private static bool IsFileLocked(string path)
        {
            try
            {
                if (!File.Exists(path)) return false; // <- don't try to open non-existing files
                using var _ = File.Open(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
                return false;
            }
            catch (IOException) { return true; }
            catch { return true; }
        }
        /// <summary>
        /// Generates a conflict-free output path by probing numbered/timestamped variants.
        /// </summary>
        private static string GetSafeOutputPath(string desiredPath)
        {
            // If there is no conflict, take it.
            if (!File.Exists(desiredPath) && !IsFileLocked(desiredPath))
                return desiredPath;
            // Try numbered variants: filename (1).ext, (2), ...
            var dir = Path.GetDirectoryName(desiredPath) ?? ".";
            var name = Path.GetFileNameWithoutExtension(desiredPath);
            var ext = Path.GetExtension(desiredPath);
            for (int i = 1; i <= 99; i++)
            {
                var candidate = Path.Combine(dir, $"{name} ({i}){ext}");
                if (!File.Exists(candidate) && !IsFileLocked(candidate))
                    return candidate;
            }
            // As a last resort, timestamped fallback
            var stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            var tsPath = Path.Combine(dir, $"{name}_{stamp}{ext}");
            return tsPath;
        }
        /// <summary>
        /// Reads the raw CSV into dynamic rows while detecting delimiter and optional columns.
        /// </summary>
        private static DynamicRow[] ReadCsv(string path, out bool hasVisitorType)
        {
            using var fs = File.OpenRead(path);
            using var sr = new StreamReader(fs, DetectEncoding(fs) ?? new UTF8Encoding(false), detectEncodingFromByteOrderMarks: true);
            string? sample = sr.ReadLine();
            if (sample == null) throw new InvalidOperationException("Empty CSV.");
            char delimiter = sample.Contains(';') && !sample.Contains(',') ? ';' : ',';
            sr.DiscardBufferedData(); fs.Seek(0, SeekOrigin.Begin);
            var config = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                Delimiter = delimiter.ToString(),
                IgnoreBlankLines = true,
                DetectColumnCountChanges = false,
                BadDataFound = null,
                TrimOptions = TrimOptions.Trim,
                MissingFieldFound = null
            };
            var list = new List<DynamicRow>();
            using var csv = new CsvReader(sr, config);
            csv.Read();
            csv.ReadHeader();
            var headers = (csv.HeaderRecord ?? Array.Empty<string>()).ToList();
            hasVisitorType = headers.Any(h => string.Equals(h, "Visitor Type", StringComparison.OrdinalIgnoreCase));
            while (csv.Read())
            {
                var dict = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);
                foreach (var h in headers)
                {
                    string? v = null;
                    try { v = csv.GetField(h); } catch { }
                    dict[h] = string.IsNullOrWhiteSpace(v) ? null : v;
                }
                list.Add(new DynamicRow(dict));
            }
            return list.ToArray();
        }
        /// <summary>
        /// Peeks the stream for BOM markers to determine encoding hints.
        /// </summary>
        private static Encoding? DetectEncoding(Stream s)
        {
            long pos = s.Position;
            Span<byte> bom = stackalloc byte[4];
            int read = s.Read(bom);
            s.Seek(pos, SeekOrigin.Begin);
            if (read >= 3 && bom[0] == 0xEF && bom[1] == 0xBB && bom[2] == 0xBF) return new UTF8Encoding(true);
            if (read >= 2 && bom[0] == 0xFF && bom[1] == 0xFE) return Encoding.Unicode;
            if (read >= 2 && bom[0] == 0xFE && bom[1] == 0xFF) return Encoding.BigEndianUnicode;
            return null;
        }
        /// <summary>
        /// Attempts to parse a Date of Birth from many common formats.
        /// </summary>
        private static bool TryParseDob(string? s, out DateTime dob)
        {
            dob = default;
            if (string.IsNullOrWhiteSpace(s)) return false;
            string[] fmts = {
                "yyyy-MM-dd","dd/MM/yyyy","MM/dd/yyyy","yyyy/MM/dd","dd-MM-yyyy","MM-dd-yyyy",
                "d/M/yyyy","M/d/yyyy","dd.MM.yyyy","d.M.yyyy","d-MMM-yy","d MMM yyyy"
            };
            if (DateTime.TryParseExact(s, fmts, CultureInfo.InvariantCulture, DateTimeStyles.None, out dob)) return true;
            if (DateTime.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.None, out dob)) return true;
            if (DateTime.TryParse(s, CultureInfo.CurrentCulture, DateTimeStyles.None, out dob)) return true;
            return false;
        }
        /// <summary>
        /// Maps a birth date into the requested window and returns the next celebration date if applicable.
        /// </summary>
        private static DateTime? NextBirthdayInWindow(DateTime dob, DateTime start, DateTime end)
        {
            static DateTime MapToYear(DateTime birth, int year)
            {
                int day = birth.Day; int month = birth.Month;
                if (month == 2 && day == 29 && !DateTime.IsLeapYear(year)) day = 28;
                return new DateTime(year, month, day);
            }
            var s = start.Date; var e = end.Date;
            var cand1 = MapToYear(dob, s.Year);
            if (e >= s)
            {
                if (cand1 >= s && cand1 <= e) return cand1;
                return null;
            }
            if (cand1 >= s) return cand1; // tail of year
            var cand2 = MapToYear(dob, s.Year + 1);
            if (cand2 <= e.AddYears(1)) return cand2;
            return null;
        }
        /// <summary>
        /// Calculates a person's age on the specified reference date.
        /// </summary>
        private static int AgeOnDate(DateTime dob, DateTime refDate)
        {
            int age = refDate.Year - dob.Year;
            if (refDate.Month < dob.Month || (refDate.Month == dob.Month && refDate.Day < dob.Day)) age--;
            return age;
        }
        /// <summary>
        /// Chooses the most suitable guardian row for a child, preferring residents when available.
        /// </summary>
        private static Row? ChooseGuardian(List<Row> adults, bool preferResident)
        {
            IEnumerable<Row> pool = adults;
            if (preferResident)
            {
                var residents = adults.Where(a => string.Equals(a.VisitorType, "Resident", StringComparison.OrdinalIgnoreCase)).ToList();
                if (residents.Count > 0) pool = residents;
            }
            return pool.OrderByDescending(a => !string.IsNullOrWhiteSpace(a.LastName)).FirstOrDefault();
        }
        /// <summary>
        /// Writes the filtered results to CSV, adding headers dynamically based on source columns.
        /// </summary>
        private static string WriteCsv(List<Output> rows, string path, bool hasVisitorType)
        {
            var savePath = GetSafeOutputPath(path);
            using var fs = File.Create(savePath);
            using var sw = new StreamWriter(fs, new UTF8Encoding(false));
            using var csv = new CsvWriter(sw, CultureInfo.InvariantCulture);
            var headers = new List<string> { "First Name","Last Name","Email","Mobile Number","Mobile (Normalized)","Date of Birth" };
            if (hasVisitorType) headers.Add("Visitor Type");
            headers.AddRange(new[] { "Parent Name","Age","Birthday day","Birthday month" });
            foreach (var h in headers) csv.WriteField(h);
            csv.NextRecord();
            foreach (var r in rows)
            {
                csv.WriteField(r.FirstName);
                csv.WriteField(r.LastName);
                csv.WriteField(r.Email);
                csv.WriteField(r.Mobile);
                csv.WriteField(r.NormalizedMobile);   // NEW
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
        /// Writes the filtered results to an XLSX workbook with table formatting applied.
        /// </summary>
        private static string WriteXlsx(List<Output> rows, string path, bool hasVisitorType)
        {
            var savePath = GetSafeOutputPath(path);
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Birthdays");
            var headers = new List<string> { "First Name","Last Name","Email","Mobile Number","Mobile (Normalized)","Date of Birth" };
            if (hasVisitorType) headers.Add("Visitor Type");
            headers.AddRange(new[] { "Parent Name","Age","Birthday day","Birthday month" });
            for (int c = 0; c < headers.Count; c++) ws.Cell(1, c + 1).Value = headers[c];
            for (int r = 0; r < rows.Count; r++)
            {
                var row = rows[r]; int c = 1;
                ws.Cell(r + 2, c++).Value = row.FirstName ?? string.Empty;
                ws.Cell(r + 2, c++).Value = row.LastName ?? string.Empty;
                ws.Cell(r + 2, c++).Value = row.Email ?? string.Empty;
                ws.Cell(r + 2, c++).Value = row.Mobile ?? string.Empty;
                ws.Cell(r + 2, c++).Value = row.NormalizedMobile ?? string.Empty;  // NEW
                ws.Cell(r + 2, c++).Value = row.DateOfBirth ?? string.Empty;
                if (hasVisitorType) ws.Cell(r + 2, c++).Value = row.VisitorType ?? string.Empty;
                ws.Cell(r + 2, c++).Value = row.ParentName ?? string.Empty;
                ws.Cell(r + 2, c++).Value = row.Age;
                ws.Cell(r + 2, c++).Value = row.BirthdayDay;
                ws.Cell(r + 2, c++).Value = row.BirthdayMonth;
            }
            var lastRow = Math.Max(1, rows.Count + 1);
            var lastCol = headers.Count;
            var range = ws.Range(1, 1, lastRow, lastCol);
            var table = range.CreateTable("BirthdaysTable");
            table.Theme = XLTableTheme.TableStyleMedium2;
            if (rows.Count > 0)
            {
                var ageCol = headers.IndexOf("Age") + 1;
                var dayCol = headers.IndexOf("Birthday day") + 1;
                var monthCol = headers.IndexOf("Birthday month") + 1;
                if (ageCol > 0) ws.Column(ageCol).Style.NumberFormat.Format = "0";
                if (dayCol > 0) ws.Column(dayCol).Style.NumberFormat.Format = "0";
                if (monthCol > 0) ws.Column(monthCol).Style.NumberFormat.Format = "0";
            }
            ws.Columns().AdjustToContents();
            wb.SaveAs(savePath);
            return savePath;
        }
        /// <summary>
        /// Helper for safely reporting progress increments back to the caller.
        /// </summary>
        private static void ReportProgress(ProcOptions o, int pct) => o.Progress?.Report(pct);
    }
}
