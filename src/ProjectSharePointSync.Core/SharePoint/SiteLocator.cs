using System;
using System.Collections.Concurrent;
using System.Threading;
using System.Threading.Tasks;

namespace ProjectSharePointSync.Core.SharePoint
{
    /// <summary>
    /// Resolves a SharePoint site URL to a Microsoft Graph site id.
    /// </summary>
    public sealed class SiteLocator
    {
        private readonly GraphClient _graph;
        private readonly ConcurrentDictionary<string, SiteInfo> _cache = new ConcurrentDictionary<string, SiteInfo>(StringComparer.OrdinalIgnoreCase);

        public SiteLocator(GraphClient graph)
        {
            _graph = graph;
        }

        public async Task<SiteInfo> ResolveAsync(string siteUrl, CancellationToken ct)
        {
            SiteInfo cached;
            var normalized = Normalize(siteUrl);
            if (_cache.TryGetValue(normalized, out cached)) return cached;

            var uri = new Uri(normalized);
            var path = uri.AbsolutePath.TrimEnd('/');
            var relative = string.IsNullOrEmpty(path) || path == "/"
                ? "/sites/" + uri.Host
                : "/sites/" + uri.Host + ":" + path;

            var json = await _graph.GetAsync(relative + "?$select=id,displayName,webUrl", ct).ConfigureAwait(false);
            var info = new SiteInfo((string)json["id"], (string)json["displayName"], (string)json["webUrl"]);
            _cache[normalized] = info;
            return info;
        }

        public void Invalidate(string siteUrl)
        {
            SiteInfo removed;
            _cache.TryRemove(Normalize(siteUrl), out removed);
        }

        private static string Normalize(string siteUrl)
        {
            Uri uri;
            if (!Uri.TryCreate((siteUrl ?? string.Empty).Trim(), UriKind.Absolute, out uri) || uri.Scheme != Uri.UriSchemeHttps)
            {
                throw new ArgumentException("Invalid SharePoint site URL: '" + siteUrl + "'. Expected e.g. https://contoso.sharepoint.com/sites/pmo");
            }

            return uri.GetLeftPart(UriPartial.Path).TrimEnd('/');
        }
    }

    public sealed class SiteInfo
    {
        public SiteInfo(string id, string displayName, string webUrl)
        {
            Id = id;
            DisplayName = displayName;
            WebUrl = webUrl;
        }

        public string Id { get; private set; }

        public string DisplayName { get; private set; }

        public string WebUrl { get; private set; }
    }
}
