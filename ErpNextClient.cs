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

namespace BirthdayExtractor
{
    /// <summary>
    /// Lightweight REST client for interacting with ERPNext leads.
    /// Handles deduplication queries and lead creation payloads.
    /// </summary>
    internal sealed class ErpNextClient : IDisposable
    {
        private readonly HttpClient _http;
        private static readonly JsonSerializerOptions _jsonOptions = new()
        {
            DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase
        };

        public ErpNextClient(string baseUrl, string apiKey, string apiSecret)
        {
            if (string.IsNullOrWhiteSpace(baseUrl)) throw new ArgumentException("Base URL is required", nameof(baseUrl));
            if (string.IsNullOrWhiteSpace(apiKey)) throw new ArgumentException("API key is required", nameof(apiKey));
            if (string.IsNullOrWhiteSpace(apiSecret)) throw new ArgumentException("API secret is required", nameof(apiSecret));

            if (!baseUrl.EndsWith('/')) baseUrl += "/";
            _http = new HttpClient
            {
                BaseAddress = new Uri(baseUrl, UriKind.Absolute)
            };
            _http.DefaultRequestHeaders.Accept.Clear();
            _http.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            _http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("token", $"{apiKey}:{apiSecret}");
            _http.DefaultRequestHeaders.UserAgent.ParseAdd($"BirthdayExtractor/{AppVersion.Display}");
        }

        public async Task<HashSet<string>> FetchExistingKeysAsync(IEnumerable<string> keys, CancellationToken cancellationToken)
        {
            var comparer = StringComparer.OrdinalIgnoreCase;
            var distinct = keys.Where(k => !string.IsNullOrWhiteSpace(k)).Distinct(comparer).ToList();
            var existing = new HashSet<string>(comparer);
            if (distinct.Count == 0)
            {
                return existing;
            }

            const int batchSize = 50;
            for (var offset = 0; offset < distinct.Count; offset += batchSize)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var slice = distinct.Skip(offset).Take(batchSize).ToList();
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

        public async Task CreateLeadAsync(ExtractedLead lead, DateTime now, CancellationToken cancellationToken)
        {
            var payload = BuildLeadPayload(lead, now);
            var json = JsonSerializer.Serialize(payload, _jsonOptions);
            using var content = new StringContent(json, Encoding.UTF8, "application/json");
            using var resp = await _http.PostAsync("api/resource/Lead", content, cancellationToken).ConfigureAwait(false);
            var body = await resp.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
            if (!resp.IsSuccessStatusCode)
            {
                throw new InvalidOperationException($"ERPNext create failed ({(int)resp.StatusCode}): {body}");
            }
        }

        internal static IReadOnlyList<string> GetMissingRequiredFields(ExtractedLead lead)
        {
            var missing = new List<string>();
            var (parentFirst, parentLast) = ResolveGuardianNames(lead);
            if (string.IsNullOrWhiteSpace(parentFirst) && string.IsNullOrWhiteSpace(parentLast))
            {
                missing.Add("first_name/last_name");
            }

            var phone = string.IsNullOrWhiteSpace(lead.NormalizedMobile) ? lead.Mobile : lead.NormalizedMobile;
            if (string.IsNullOrWhiteSpace(phone))
            {
                missing.Add("mobile_no");
            }

            if (string.IsNullOrWhiteSpace(lead.Email))
            {
                missing.Add("custom_email");
            }

            var childName = string.Join(" ", new[] { lead.ChildFirstName, lead.ChildLastName }
                .Where(s => !string.IsNullOrWhiteSpace(s))
                .Select(s => s!.Trim()));
            if (string.IsNullOrWhiteSpace(childName))
            {
                missing.Add("custom_child_names");
            }

            if (lead.Age <= 0)
            {
                missing.Add("custom_child_ages");
            }

            if (string.IsNullOrWhiteSpace(lead.DateOfBirth))
            {
                missing.Add("child_date_of_birth");
            }

            return missing;
        }

        private static Dictionary<string, object?> BuildLeadPayload(ExtractedLead lead, DateTime now)
        {
            var (parentFirst, parentLast) = ResolveGuardianNames(lead);
            var childName = string.Join(" ", new[] { lead.ChildFirstName, lead.ChildLastName }
                .Where(s => !string.IsNullOrWhiteSpace(s))
                .Select(s => s!.Trim()));
            var phone = string.IsNullOrWhiteSpace(lead.NormalizedMobile) ? lead.Mobile : lead.NormalizedMobile;
            var notesBuilder = new StringBuilder();
            notesBuilder.Append("added by automation on ");
            notesBuilder.Append(now.ToString("yyyy-MM-dd HH:mm"));
            notesBuilder.AppendLine();
            notesBuilder.Append("Child's DOB: ");
            notesBuilder.Append(string.IsNullOrWhiteSpace(lead.DateOfBirth) ? "Unknown" : lead.DateOfBirth);

            var payload = new Dictionary<string, object?>
            {
                ["doctype"] = "Lead",
                ["naming_series"] = "CRM-LEAD-.YYYY.-",
                ["source"] = "Outbound",
                ["status"] = string.Empty,
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

        private static (string? First, string? Last) ResolveGuardianNames(ExtractedLead lead)
        {
            var first = lead.ParentFirstName;
            var last = lead.ParentLastName;
            if (!string.IsNullOrWhiteSpace(first) || !string.IsNullOrWhiteSpace(last))
            {
                return (Normalize(first), Normalize(last));
            }

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
                return (Normalize(parts[0]), null);
            }

            var potentialLast = parts[^1];
            var potentialFirst = string.Join(" ", parts.Take(parts.Length - 1));
            return (Normalize(potentialFirst), Normalize(potentialLast));
        }

        private static string? Normalize(string? input)
            => string.IsNullOrWhiteSpace(input) ? null : input.Trim();

        public void Dispose()
        {
            _http.Dispose();
        }
    }
}
