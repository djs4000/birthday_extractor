using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace BirthdayExtractor
{
    /// <summary>
    /// Immutable options used when uploading extracted leads to ERPNext.
    /// </summary>
    internal sealed class ErpNextUploadOptions
    {
        public ErpNextUploadOptions(string baseUrl, string apiKey, string apiSecret)
        {
            BaseUrl = baseUrl ?? throw new ArgumentNullException(nameof(baseUrl));
            ApiKey = apiKey ?? throw new ArgumentNullException(nameof(apiKey));
            ApiSecret = apiSecret ?? throw new ArgumentNullException(nameof(apiSecret));
        }

        public string BaseUrl { get; }
        public string ApiKey { get; }
        public string ApiSecret { get; }

        /// <summary>
        /// Optional timestamp injected into the ERPNext notes payload for repeatable CLI runs.
        /// When null, <see cref="DateTime.Now"/> will be used.
        /// </summary>
        public DateTime? UploadTimestamp { get; set; }
    }

    /// <summary>
    /// Summary data describing the outcome of an ERPNext upload attempt.
    /// </summary>
    internal sealed class ErpNextUploadSummary
    {
        public int TotalLeads { get; init; }
        public int MissingBusinessKey { get; init; }
        public int MissingRequiredFields { get; init; }
        public int Candidates { get; init; }
        public int Duplicates { get; init; }
        public int Created { get; init; }
        public int Failed { get; init; }
    }

    /// <summary>
    /// Shared helper that encapsulates the ERPNext upload workflow so it can be reused by the UI and CLI paths.
    /// </summary>
    internal static class ErpNextUploader
    {
        public static async Task<ErpNextUploadSummary> UploadAsync(
            IEnumerable<ExtractedLead> leads,
            ErpNextUploadOptions options,
            Action<string> log,
            CancellationToken cancellationToken)
        {
            if (leads is null) throw new ArgumentNullException(nameof(leads));
            if (options is null) throw new ArgumentNullException(nameof(options));
            if (log is null) throw new ArgumentNullException(nameof(log));

            var allLeads = leads.ToList();
            var withKeys = allLeads.Where(l => !string.IsNullOrWhiteSpace(l.BusinessKey)).ToList();
            var missingKeyCount = allLeads.Count - withKeys.Count;
            if (missingKeyCount > 0)
            {
                log($"WARN: {missingKeyCount} lead(s) are missing a business key and will be skipped.");
            }

            var uploadCandidates = new List<ExtractedLead>(withKeys.Count);
            int missingFieldSkips = 0;
            foreach (var lead in withKeys)
            {
                var missingFields = ErpNextClient.GetMissingRequiredFields(lead);
                if (missingFields.Count > 0)
                {
                    missingFieldSkips++;
                    log($"Skipping {lead.BusinessKey}: missing {string.Join(", ", missingFields)}.");
                    continue;
                }

                uploadCandidates.Add(lead);
            }

            if (missingFieldSkips > 0)
            {
                log($"Skipped {missingFieldSkips} lead(s) due to missing required fields.");
            }

            if (uploadCandidates.Count == 0)
            {
                log("No leads with all required fields available for upload.");
                return new ErpNextUploadSummary
                {
                    TotalLeads = allLeads.Count,
                    MissingBusinessKey = missingKeyCount,
                    MissingRequiredFields = missingFieldSkips,
                    Candidates = 0,
                    Duplicates = 0,
                    Created = 0,
                    Failed = 0
                };
            }

            using var client = new ErpNextClient(options.BaseUrl, options.ApiKey, options.ApiSecret);
            var uniqueKeys = new HashSet<string>(uploadCandidates.Select(l => l.BusinessKey!), StringComparer.OrdinalIgnoreCase);
            log($"Collected {uniqueKeys.Count} unique business key(s) from this run.");

            var existing = await client.FetchExistingKeysAsync(uniqueKeys, cancellationToken).ConfigureAwait(false);
            log($"ERPNext already contains {existing.Count} matching lead(s).");

            var toCreate = uploadCandidates.Where(l => !existing.Contains(l.BusinessKey!)).ToList();
            if (toCreate.Count == 0)
            {
                log("All leads already exist in ERPNext. Nothing to upload.");
                return new ErpNextUploadSummary
                {
                    TotalLeads = allLeads.Count,
                    MissingBusinessKey = missingKeyCount,
                    MissingRequiredFields = missingFieldSkips,
                    Candidates = uploadCandidates.Count,
                    Duplicates = existing.Count,
                    Created = 0,
                    Failed = 0
                };
            }

            int success = 0, failed = 0, index = 0;
            var timestamp = options.UploadTimestamp ?? DateTime.Now;
            foreach (var lead in toCreate)
            {
                cancellationToken.ThrowIfCancellationRequested();
                index++;
                try
                {
                    await client.CreateLeadAsync(lead, timestamp, cancellationToken).ConfigureAwait(false);
                    success++;
                    var childDisplay = string.Join(" ", new[] { lead.ChildFirstName, lead.ChildLastName }
                        .Where(s => !string.IsNullOrWhiteSpace(s)));
                    if (string.IsNullOrWhiteSpace(childDisplay)) childDisplay = lead.BusinessKey;
                    log($"Uploaded {index}/{toCreate.Count}: {childDisplay}");
                }
                catch (Exception ex)
                {
                    failed++;
                    log($"ERROR uploading {lead.BusinessKey}: {ex.Message}");
                    if (!LogRouter.IsRegisteredLogger(log))
                    {
                        LogRouter.LogException(ex, $"ERROR uploading {lead.BusinessKey}");
                    }
                }
            }

            log($"Upload complete. Created {success} lead(s), skipped {existing.Count} duplicate(s), failed {failed}.");

            return new ErpNextUploadSummary
            {
                TotalLeads = allLeads.Count,
                MissingBusinessKey = missingKeyCount,
                MissingRequiredFields = missingFieldSkips,
                Candidates = uploadCandidates.Count,
                Duplicates = existing.Count,
                Created = success,
                Failed = failed
            };
        }
    }
}
