// Namespace for the entire Birthday Extractor application.
namespace BirthdayExtractor
{
    // Imports for handling various functionalities like collections, networking, JSON serialization, and asynchronous operations.
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Net.Http;
    using System.Net.Http.Headers;
    using System.Text;
    using System.Text.Json;
    using System.Text.Json.Serialization;
    using System.Threading;
    using System.Threading.Tasks;

    /// <summary>
    /// A lightweight REST client for interacting with the ERPNext "Lead" DocType.
    /// This class handles fetching existing leads to prevent duplicates and creating new ones.
    /// It is designed to be instantiated once per upload session.
    /// </summary>
    internal sealed class ErpNextClient : IDisposable
    {
        private readonly HttpClient _http;
        private static readonly JsonSerializerOptions _jsonOptions = new()
        {
            // Configure JSON serialization to ignore null values and use camelCase naming.
            DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase
        };

        /// <summary>
        /// Initializes a new instance of the <see cref="ErpNextClient"/> class.
        /// </summary>
        /// <param name="baseUrl">The base URL of the ERPNext instance.</param>
        /// <param name="apiKey">The API key for authentication.</param>
        /// <param name="apiSecret">The API secret for authentication.</param>
        public ErpNextClient(string baseUrl, string apiKey, string apiSecret)
        {
            // Validate required parameters.
            if (string.IsNullOrWhiteSpace(baseUrl)) throw new ArgumentException("Base URL is required", nameof(baseUrl));
            if (string.IsNullOrWhiteSpace(apiKey)) throw new ArgumentException("API key is required", nameof(apiKey));
            if (string.IsNullOrWhiteSpace(apiSecret)) throw new ArgumentException("API secret is required", nameof(apiSecret));

            // Ensure the base URL ends with a slash for proper URI construction.
            if (!baseUrl.EndsWith('/')) baseUrl += "/";

            // Configure the HttpClient instance.
            _http = new HttpClient
            {
                BaseAddress = new Uri(baseUrl, UriKind.Absolute)
            };
            _http.DefaultRequestHeaders.Accept.Clear();
            _http.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            // Set the authorization header using the ERPNext token format.
            _http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("token", $"{apiKey}:{apiSecret}");
            _http.DefaultRequestHeaders.UserAgent.ParseAdd($"BirthdayExtractor/{AppVersion.Display}");
        }

        /// <summary>
        /// Fetches existing lead keys from ERPNext in batches to check for duplicates before uploading.
        /// </summary>
        /// <param name="keys">A collection of business keys to check.</param>
        /// <param name="cancellationToken">A token to cancel the operation.</param>
        /// <returns>A HashSet containing the keys that already exist in ERPNext.</returns>
        public async Task<HashSet<string>> FetchExistingKeysAsync(IEnumerable<string> keys, CancellationToken cancellationToken)
        {
            var comparer = StringComparer.OrdinalIgnoreCase;
            var distinct = keys.Where(k => !string.IsNullOrWhiteSpace(k)).Distinct(comparer).ToList();
            var existing = new HashSet<string>(comparer);
            if (distinct.Count == 0)
            {
                return existing;
            }

            // Process keys in batches to avoid creating URLs that are too long.
            const int batchSize = 50;
            // TODO: Make the batch size configurable or determine it dynamically based on URL length limits.
            for (var offset = 0; offset < distinct.Count; offset += batchSize)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var slice = distinct.Skip(offset).Take(batchSize).ToList();

                // Construct the ERPNext GET request with filters.
                var fieldsParam = JsonSerializer.Serialize(new[] { "name", "custom_birthday_key" });
                var filtersParam = JsonSerializer.Serialize(new object[]
                {
                    new object[] { "Lead", "custom_birthday_key", "in", slice }
                });
                var url = $"api/resource/Lead?fields={Uri.EscapeDataString(fieldsParam)}&filters={Uri.EscapeDataString(filtersParam)}&limit_page_length={slice.Count}";

                using var resp = await _http.GetAsync(url, cancellationToken).ConfigureAwait(false);
                var body = await resp.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
                if (!resp.IsSuccessStatusCode)
                {
                    throw new InvalidOperationException($"ERPNext lookup failed ({(int)resp.StatusCode}): {body}");
                }

                // Parse the JSON response to extract the existing keys.
                using var doc = JsonDocument.Parse(body);
                if (!doc.RootElement.TryGetProperty("data", out var dataArray))
                {
                    continue;
                }

                foreach (var item in dataArray.EnumerateArray())
                {
                    if (item.TryGetProperty("custom_birthday_key", out var keyElement))
                    {
                        var key = keyElement.GetString();
                        if (!string.IsNullOrWhiteSpace(key)) existing.Add(key);
                    }
                }
            }

            return existing;
        }

        /// <summary>
        /// Creates a new Lead in ERPNext based on the extracted lead information.
        /// </summary>
        public async Task CreateLeadAsync(ExtractedLead lead, DateTime now, CancellationToken cancellationToken)
        {
            var payload = BuildLeadPayload(lead, now);
            var json = JsonSerializer.Serialize(payload, _jsonOptions);
            using var content = new StringContent(json, Encoding.UTF8, "application/json");

            // Send the POST request to create the lead.
            using var resp = await _http.PostAsync("api/resource/Lead", content, cancellationToken).ConfigureAwait(false);
            var body = await resp.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
            if (!resp.IsSuccessStatusCode)
            {
                throw new InvalidOperationException($"ERPNext create failed ({(int)resp.StatusCode}): {body}");
            }
        }

        /// <summary>
        /// Checks an extracted lead for missing fields that are required for creating a valid ERPNext lead.
        /// </summary>
        /// <returns>A list of missing required field names.</returns>
        internal static IReadOnlyList<string> GetMissingRequiredFields(ExtractedLead lead)
        {
            var missing = new List<string>();
            var (parentFirst, parentLast) = ResolveGuardianNames(lead);

            // Guardian name is required.
            if (string.IsNullOrWhiteSpace(parentFirst) && string.IsNullOrWhiteSpace(parentLast))
            {
                missing.Add("first_name/last_name");
            }

            // A phone number is required.
            var phone = string.IsNullOrWhiteSpace(lead.NormalizedMobile) ? lead.Mobile : lead.NormalizedMobile;
            if (string.IsNullOrWhiteSpace(phone))
            {
                missing.Add("mobile_no");
            }

            // An email is required.
            if (string.IsNullOrWhiteSpace(lead.Email))
            {
                missing.Add("custom_email");
            }

            // Child name is required.
            var childName = string.Join(" ", new[] { lead.ChildFirstName, lead.ChildLastName }
                .Where(s => !string.IsNullOrWhiteSpace(s))
                .Select(s => s!.Trim()));
            if (string.IsNullOrWhiteSpace(childName))
            {
                missing.Add("custom_child_names");
            }

            // Child age is required.
            if (lead.Age <= 0)
            {
                missing.Add("custom_child_ages");
            }

            // Child date of birth is required.
            if (string.IsNullOrWhiteSpace(lead.DateOfBirth))
            {
                missing.Add("child_date_of_birth");
            }

            return missing;
        }

        /// <summary>
        /// Constructs the JSON payload for creating a new ERPNext Lead from an ExtractedLead object.
        /// </summary>
        private static Dictionary<string, object?> BuildLeadPayload(ExtractedLead lead, DateTime now)
        {
            var (parentFirst, parentLast) = ResolveGuardianNames(lead);
            var childName = string.Join(" ", new[] { lead.ChildFirstName, lead.ChildLastName }
                .Where(s => !string.IsNullOrWhiteSpace(s))
                .Select(s => s!.Trim()));
            var phone = string.IsNullOrWhiteSpace(lead.NormalizedMobile) ? lead.Mobile : lead.NormalizedMobile;

            // Build a notes string with metadata.
            var notesBuilder = new StringBuilder();
            notesBuilder.Append("Added by Birthday Extractor on ");
            notesBuilder.Append(now.ToString("yyyy-MM-dd HH:mm"));
            notesBuilder.AppendLine();
            notesBuilder.Append("Child's DOB: ");
            notesBuilder.Append(string.IsNullOrWhiteSpace(lead.DateOfBirth) ? "Unknown" : lead.DateOfBirth);

            // Map the extracted lead data to the ERPNext Lead DocType fields.
            var payload = new Dictionary<string, object?>
            {
                ["doctype"] = "Lead",
                ["naming_series"] = "CRM-LEAD-.YYYY.-",
                ["source"] = "Outbound", // TODO: Make this configurable.
                ["status"] = "Initial Contact", // TODO: Make this configurable.
                ["first_name"] = parentFirst,
                ["last_name"] = parentLast,
                ["mobile_no"] = phone,
                ["whatsapp_no"] = phone,
                ["custom_email"] = lead.Email,
                ["custom_booking_type"] = "Birthday Party",
                ["custom_next_follow_up"] = now.ToString("yyyy-MM-dd"),
                ["custom_child_names"] = string.IsNullOrWhiteSpace(childName) ? null : childName,
                ["custom_child_ages"] = lead.Age,
                ["custom_birthday_addons_and_notes"] = notesBuilder.ToString(),
                ["custom_birthday_key"] = string.IsNullOrWhiteSpace(lead.BusinessKey) ? null : lead.BusinessKey
            };

            return payload;
        }

        /// <summary>
        /// Resolves the guardian's first and last names from the available parent name fields.
        /// It prioritizes the separate first/last name fields and falls back to splitting the full name.
        /// </summary>
        private static (string? First, string? Last) ResolveGuardianNames(ExtractedLead lead)
        {
            var first = lead.ParentFirstName;
            var last = lead.ParentLastName;
            // If separate first/last names are provided, use them.
            if (!string.IsNullOrWhiteSpace(first) || !string.IsNullOrWhiteSpace(last))
            {
                return (Normalize(first), Normalize(last));
            }

            // Otherwise, try to parse the full ParentName.
            if (string.IsNullOrWhiteSpace(lead.ParentName))
            {
                return (null, null);
            }

            var parts = lead.ParentName.Split(' ', StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 0)
            {
                return (null, null);
            }
            if (parts.Length == 1)
            {
                // Assume a single part is the first name.
                return (Normalize(parts[0]), null);
            }

            // Assume the last part is the last name and the rest is the first name.
            var potentialLast = parts[^1];
            var potentialFirst = string.Join(" ", parts.Take(parts.Length - 1));
            return (Normalize(potentialFirst), Normalize(potentialLast));
        }

        /// <summary>
        /// A simple helper to trim a string or return null if it's whitespace.
        /// </summary>
        private static string? Normalize(string? input)
            => string.IsNullOrWhiteSpace(input) ? null : input.Trim();

        /// <summary>
        /// Disposes the underlying HttpClient instance.
        /// </summary>
        public void Dispose()
        {
            _http.Dispose();
        }
    }
}