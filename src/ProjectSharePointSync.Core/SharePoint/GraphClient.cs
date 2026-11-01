using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using ProjectSharePointSync.Core.Diagnostics;

namespace ProjectSharePointSync.Core.SharePoint
{
    /// <summary>
    /// Thin Microsoft Graph REST client with throttling-aware retries and JSON batching.
    /// </summary>
    public sealed class GraphClient : IDisposable
    {
        public const string BaseUrl = "https://graph.microsoft.com/v1.0";

        /// <summary>Maximum number of requests in one Graph JSON batch.</summary>
        public const int MaxBatchSize = 20;

        private const int MaxAttempts = 6;

        private readonly HttpClient _http;
        private readonly Func<CancellationToken, Task<string>> _tokenFactory;

        static GraphClient()
        {
            ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
        }

        public GraphClient(Func<CancellationToken, Task<string>> tokenFactory)
        {
            if (tokenFactory == null) throw new ArgumentNullException("tokenFactory");
            _tokenFactory = tokenFactory;
            _http = new HttpClient(new HttpClientHandler { UseDefaultCredentials = false, UseProxy = true })
            {
                Timeout = TimeSpan.FromMinutes(2)
            };
        }

        public Task<JObject> GetAsync(string relativeUrl, CancellationToken ct)
        {
            return SendAsync(HttpMethod.Get, relativeUrl, null, ct);
        }

        public Task<JObject> PostAsync(string relativeUrl, JObject body, CancellationToken ct)
        {
            return SendAsync(HttpMethod.Post, relativeUrl, body, ct);
        }

        /// <summary>Follows @odata.nextLink and returns all items of the "value" collection.</summary>
        public async Task<List<JObject>> GetAllAsync(string relativeUrl, CancellationToken ct)
        {
            var all = new List<JObject>();
            string url = relativeUrl;
            while (url != null)
            {
                var page = await SendAsync(HttpMethod.Get, url, null, ct).ConfigureAwait(false);
                var values = page["value"] as JArray;
                if (values != null)
                {
                    all.AddRange(values.OfType<JObject>());
                }

                var next = (string)page["@odata.nextLink"];
                url = string.IsNullOrEmpty(next) ? null : next;
            }

            return all;
        }

        public async Task<JObject> SendAsync(HttpMethod method, string url, JObject body, CancellationToken ct)
        {
            var absolute = url.StartsWith("https://", StringComparison.OrdinalIgnoreCase) ? url : BaseUrl + url;

            for (var attempt = 1; ; attempt++)
            {
                ct.ThrowIfCancellationRequested();
                using (var request = new HttpRequestMessage(method, absolute))
                {
                    var token = await _tokenFactory(ct).ConfigureAwait(false);
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                    request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                    request.Headers.TryAddWithoutValidation("Prefer", "HonorNonIndexedQueriesWarningMayFailRandomly");
                    if (body != null)
                    {
                        request.Content = new StringContent(body.ToString(Newtonsoft.Json.Formatting.None), Encoding.UTF8, "application/json");
                    }

                    HttpResponseMessage response;
                    try
                    {
                        response = await _http.SendAsync(request, ct).ConfigureAwait(false);
                    }
                    catch (HttpRequestException ex)
                    {
                        if (attempt >= MaxAttempts) throw;
                        Log.Warn("Network error calling Graph (attempt " + attempt + "): " + ex.Message);
                        await Task.Delay(Backoff(attempt), ct).ConfigureAwait(false);
                        continue;
                    }

                    using (response)
                    {
                        var text = response.Content == null ? null : await response.Content.ReadAsStringAsync().ConfigureAwait(false);

                        if (IsTransient(response.StatusCode) && attempt < MaxAttempts)
                        {
                            var delay = RetryAfter(response) ?? Backoff(attempt);
                            Log.Warn(string.Format("Graph returned {0} for {1} {2}, retrying in {3:0.0}s.", (int)response.StatusCode, method, url, delay.TotalSeconds));
                            await Task.Delay(delay, ct).ConfigureAwait(false);
                            continue;
                        }

                        if (!response.IsSuccessStatusCode)
                        {
                            throw ToException(response.StatusCode, text);
                        }

                        return string.IsNullOrWhiteSpace(text) ? new JObject() : JObject.Parse(text);
                    }
                }
            }
        }

        /// <summary>
        /// Executes requests using Graph JSON batching (20 per call). Requests that are
        /// throttled inside a batch are retried. Returns one response per request, in order.
        /// </summary>
        public async Task<IList<BatchResponse>> BatchAsync(IList<BatchRequest> requests, CancellationToken ct)
        {
            var results = new BatchResponse[requests.Count];
            var pending = Enumerable.Range(0, requests.Count).ToList();

            for (var attempt = 1; pending.Count > 0; attempt++)
            {
                var retry = new List<int>();
                TimeSpan maxDelay = TimeSpan.Zero;

                foreach (var chunk in Chunk(pending, MaxBatchSize))
                {
                    var payload = new JArray();
                    foreach (var index in chunk)
                    {
                        var r = requests[index];
                        var item = new JObject
                        {
                            ["id"] = index.ToString(CultureInfo.InvariantCulture),
                            ["method"] = r.Method,
                            ["url"] = r.Url
                        };
                        if (r.Body != null)
                        {
                            item["body"] = r.Body;
                            item["headers"] = new JObject { ["Content-Type"] = "application/json" };
                        }

                        payload.Add(item);
                    }

                    var response = await PostAsync("/$batch", new JObject { ["requests"] = payload }, ct).ConfigureAwait(false);
                    foreach (var resp in (response["responses"] as JArray ?? new JArray()).OfType<JObject>())
                    {
                        var index = int.Parse((string)resp["id"], CultureInfo.InvariantCulture);
                        var status = (int?)resp["status"] ?? 0;

                        if (IsTransient((HttpStatusCode)status) && attempt < MaxAttempts)
                        {
                            retry.Add(index);
                            var ra = ParseRetryAfter(resp["headers"] as JObject) ?? Backoff(attempt);
                            if (ra > maxDelay) maxDelay = ra;
                            continue;
                        }

                        results[index] = new BatchResponse(status, resp["body"] as JObject);
                    }

                    // Requests missing from the response (should not happen) are retried as well.
                    foreach (var index in chunk)
                    {
                        if (results[index] == null && !retry.Contains(index))
                        {
                            if (attempt < MaxAttempts) retry.Add(index);
                            else results[index] = new BatchResponse(0, null);
                        }
                    }
                }

                pending = retry;
                if (pending.Count > 0)
                {
                    Log.Warn(string.Format("{0} batched request(s) throttled, retrying in {1:0.0}s.", pending.Count, maxDelay.TotalSeconds));
                    await Task.Delay(maxDelay, ct).ConfigureAwait(false);
                }
            }

            return results;
        }

        public void Dispose()
        {
            _http.Dispose();
        }

        internal static GraphException ToException(HttpStatusCode status, string body)
        {
            string code = status.ToString();
            string message = body;
            try
            {
                var json = string.IsNullOrWhiteSpace(body) ? null : JObject.Parse(body);
                var error = json == null ? null : json["error"];
                if (error != null)
                {
                    code = (string)error["code"] ?? code;
                    message = (string)error["message"] ?? message;
                }
            }
            catch
            {
                // not JSON
            }

            return new GraphException(status, code, message);
        }

        private static bool IsTransient(HttpStatusCode status)
        {
            var code = (int)status;
            return code == 429 || code == 503 || code == 504;
        }

        private static TimeSpan Backoff(int attempt)
        {
            return TimeSpan.FromSeconds(Math.Min(60, Math.Pow(2, attempt)));
        }

        private static TimeSpan? RetryAfter(HttpResponseMessage response)
        {
            var ra = response.Headers.RetryAfter;
            if (ra == null) return null;
            if (ra.Delta.HasValue) return ra.Delta.Value;
            if (ra.Date.HasValue)
            {
                var delta = ra.Date.Value - DateTimeOffset.UtcNow;
                return delta > TimeSpan.Zero ? delta : TimeSpan.FromSeconds(1);
            }

            return null;
        }

        private static TimeSpan? ParseRetryAfter(JObject headers)
        {
            if (headers == null) return null;
            foreach (var prop in headers.Properties())
            {
                int seconds;
                if (string.Equals(prop.Name, "Retry-After", StringComparison.OrdinalIgnoreCase)
                    && int.TryParse((string)prop.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out seconds))
                {
                    return TimeSpan.FromSeconds(Math.Max(1, seconds));
                }
            }

            return null;
        }

        private static IEnumerable<List<int>> Chunk(List<int> source, int size)
        {
            for (var i = 0; i < source.Count; i += size)
            {
                yield return source.GetRange(i, Math.Min(size, source.Count - i));
            }
        }
    }

    public sealed class BatchRequest
    {
        public BatchRequest(string method, string url, JObject body = null)
        {
            Method = method;
            Url = url;
            Body = body;
        }

        public string Method { get; private set; }

        /// <summary>URL relative to /v1.0, e.g. /sites/{id}/lists/{id}/items</summary>
        public string Url { get; private set; }

        public JObject Body { get; private set; }
    }

    public sealed class BatchResponse
    {
        public BatchResponse(int status, JObject body)
        {
            Status = status;
            Body = body;
        }

        public int Status { get; private set; }

        public JObject Body { get; private set; }

        public bool IsSuccess
        {
            get { return Status >= 200 && Status < 300; }
        }

        public string ErrorMessage
        {
            get
            {
                var error = Body == null ? null : Body["error"];
                var message = error == null ? null : (string)error["message"];
                return string.Format(CultureInfo.InvariantCulture, "HTTP {0}: {1}", Status, message ?? "(no details)");
            }
        }
    }
}
