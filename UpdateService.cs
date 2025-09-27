using System;
using System.Globalization;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace BirthdayExtractor
{
    /// <summary>
    /// Handles polling GitHub for the latest private release and retrieving
    /// the published binary when a newer version is available.
    /// </summary>
    internal sealed class UpdateService : IDisposable
    {
        private readonly HttpClient _httpClient;
        private readonly bool _ownsClient;
        private readonly bool _hasPersonalAccessToken;
        private readonly string _apiEndpoint;

        public UpdateService(string repoOwner, string repoName, string? personalAccessToken, HttpClient? client = null)
        {
            _httpClient = client ?? CreateHttpClient(personalAccessToken);
            _ownsClient = client is null;
            _hasPersonalAccessToken = !string.IsNullOrWhiteSpace(personalAccessToken);
            _apiEndpoint = $"https://api.github.com/repos/{repoOwner}/{repoName}/releases/latest";
            _httpClient.DefaultRequestHeaders.UserAgent.ParseAdd($"BirthdayExtractor/{AppVersion.Display}");

            // When a pre-configured client is supplied we still need to ensure
            // auth headers are present if a token exists.
            if (client is not null && _hasPersonalAccessToken)
            {
                client.DefaultRequestHeaders.Authorization =
                    new AuthenticationHeaderValue("token", personalAccessToken);
            }
        }

        /// <summary>
        /// Queries GitHub for the most recent release and returns metadata when
        /// a newer version than <paramref name="currentVersion"/> is available.
        /// </summary>
        public async Task<ReleaseInfo?> CheckForNewerReleaseAsync(Version currentVersion, CancellationToken cancellationToken)
        {
            using var request = new HttpRequestMessage(HttpMethod.Get, _apiEndpoint);
            using var response = await _httpClient.SendAsync(request, cancellationToken);

            if (!response.IsSuccessStatusCode)
            {
                if (response.StatusCode is HttpStatusCode.Unauthorized or HttpStatusCode.Forbidden)
                {
                    throw new InvalidOperationException("GitHub authentication failed. Please verify the personal access token.");
                }

                var message = $"GitHub request failed: {(int)response.StatusCode} {response.ReasonPhrase}";
                throw new InvalidOperationException(message);
            }

            await using var stream = await response.Content.ReadAsStreamAsync(cancellationToken);
            using var document = await JsonDocument.ParseAsync(stream, cancellationToken: cancellationToken);
            var root = document.RootElement;

            var tag = root.GetProperty("tag_name").GetString();
            if (string.IsNullOrWhiteSpace(tag)) return null;

            var versionText = tag.Trim();
            if (versionText.StartsWith("v", true, CultureInfo.InvariantCulture))
            {
                versionText = versionText[1..];
            }

            if (!Version.TryParse(versionText, out var latestVersion))
            {
                return null;
            }

            if (latestVersion <= currentVersion)
            {
                return null;
            }

            if (!root.TryGetProperty("assets", out var assetsElement))
            {
                return null;
            }

            ReleaseAsset? selectedAsset = null;
            foreach (var asset in assetsElement.EnumerateArray())
            {
                var name = asset.GetProperty("name").GetString();
                var downloadUrl = asset.GetProperty("browser_download_url").GetString();
                var apiDownloadUrl = asset.TryGetProperty("url", out var apiUrlElement)
                    ? apiUrlElement.GetString()
                    : null;
                var size = asset.TryGetProperty("size", out var sizeElement) ? sizeElement.GetInt64() : 0L;

                if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(downloadUrl))
                {
                    continue;
                }

                var candidate = new ReleaseAsset(
                    name,
                    new Uri(downloadUrl, UriKind.Absolute),
                    !string.IsNullOrWhiteSpace(apiDownloadUrl) ? new Uri(apiDownloadUrl, UriKind.Absolute) : null,
                    size);

                // Prefer .exe payloads for the self-contained Windows app, otherwise keep the first asset.
                if (selectedAsset is null || name.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
                {
                    selectedAsset = candidate;

                    if (name.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
                    {
                        break;
                    }
                }
            }

            if (selectedAsset is null)
            {
                return null;
            }

            var releaseName = root.TryGetProperty("name", out var nameElement)
                ? nameElement.GetString() ?? tag
                : tag;

            var releaseNotes = root.TryGetProperty("body", out var bodyElement)
                ? bodyElement.GetString()
                : null;

            return new ReleaseInfo(tag, releaseName, latestVersion, releaseNotes, selectedAsset);
        }

        /// <summary>
        /// Downloads the supplied release asset to the specified destination.
        /// </summary>
        public async Task<string> DownloadAssetAsync(ReleaseAsset asset, IProgress<int>? progress, CancellationToken cancellationToken)
        {
            var destinationPath = Path.Combine(Path.GetTempPath(), asset.Name);

            var requestUri = _hasPersonalAccessToken && asset.ApiDownloadUrl is not null
                ? asset.ApiDownloadUrl
                : asset.BrowserDownloadUrl;

            using var request = new HttpRequestMessage(HttpMethod.Get, requestUri);
            request.Headers.Accept.Clear();
            request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/octet-stream"));

            using var response = await _httpClient.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, cancellationToken);

            if (!response.IsSuccessStatusCode)
            {
                if (response.StatusCode is HttpStatusCode.Unauthorized or HttpStatusCode.Forbidden)
                {
                    throw new InvalidOperationException("GitHub authentication failed while downloading the release asset.");
                }

                var message = $"Download failed: {(int)response.StatusCode} {response.ReasonPhrase}";
                throw new InvalidOperationException(message);
            }

            var totalBytes = response.Content.Headers.ContentLength;
            await using var httpStream = await response.Content.ReadAsStreamAsync(cancellationToken);
            await using var fileStream = new FileStream(destinationPath, FileMode.Create, FileAccess.Write, FileShare.None);

            var buffer = new byte[81920];
            long downloaded = 0;
            while (true)
            {
                var read = await httpStream.ReadAsync(buffer.AsMemory(0, buffer.Length), cancellationToken);
                if (read == 0) break;

                await fileStream.WriteAsync(buffer.AsMemory(0, read), cancellationToken);
                downloaded += read;

                if (totalBytes.HasValue && totalBytes.Value > 0 && progress is not null)
                {
                    var percent = (int)Math.Round(downloaded * 100d / totalBytes.Value);
                    progress.Report(Math.Max(0, Math.Min(100, percent)));
                }
            }

            progress?.Report(100);
            return destinationPath;
        }

        private static HttpClient CreateHttpClient(string? personalAccessToken)
        {
            var client = new HttpClient();
            client.DefaultRequestHeaders.UserAgent.ParseAdd($"BirthdayExtractor/{AppVersion.Display}");
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/vnd.github+json"));

            if (!string.IsNullOrWhiteSpace(personalAccessToken))
            {
                client.DefaultRequestHeaders.Authorization =
                    new AuthenticationHeaderValue("token", personalAccessToken);
            }

            return client;
        }

        public void Dispose()
        {
            if (_ownsClient)
            {
                _httpClient.Dispose();
            }
        }

        /// <summary>
        /// Describes a release returned from GitHub.
        /// </summary>
        internal sealed record ReleaseInfo(string Tag, string Title, Version Version, string? Notes, ReleaseAsset Asset);

        /// <summary>
        /// Represents a downloadable asset bundled with a release.
        /// </summary>
        internal sealed record ReleaseAsset(string Name, Uri BrowserDownloadUrl, Uri? ApiDownloadUrl, long SizeBytes);
    }
}
