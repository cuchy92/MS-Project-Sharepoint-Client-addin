using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using ProjectSharePointSync.Core.Configuration;
using ProjectSharePointSync.Core.Diagnostics;
using ProjectSharePointSync.Core.Model;
using ProjectSharePointSync.Core.Schema;

namespace ProjectSharePointSync.Core.SharePoint
{
    /// <summary>
    /// Synchronizes a <see cref="ProjectSnapshot"/> to SharePoint Online lists.
    /// Items are matched by the PsKey column; unchanged items (same PsHash) are skipped,
    /// new items are created, changed ones updated and removed ones deleted.
    /// </summary>
    public sealed class ProjectSyncService
    {
        private const int TextMaxLength = 255;
        private const int NoteMaxLength = 60000;

        private readonly GraphClient _graph;
        private readonly SiteLocator _sites;
        private readonly ListProvisioner _provisioner;

        public ProjectSyncService(GraphClient graph)
        {
            _graph = graph;
            _sites = new SiteLocator(graph);
            _provisioner = new ListProvisioner(graph);
        }

        public async Task<SiteInfo> TestConnectionAsync(string siteUrl, CancellationToken ct)
        {
            _sites.Invalidate(siteUrl);
            return await _sites.ResolveAsync(siteUrl, ct).ConfigureAwait(false);
        }

        public async Task<SyncResult> SyncAsync(ProjectSnapshot snapshot, SyncSettings settings, CancellationToken ct)
        {
            if (snapshot == null) throw new ArgumentNullException("snapshot");
            if (settings == null) throw new ArgumentNullException("settings");

            var siteUrl = string.IsNullOrWhiteSpace(snapshot.SiteUrlOverride) ? settings.SiteUrl : snapshot.SiteUrlOverride;
            var result = new SyncResult(snapshot.ProjectName) { SiteUrl = siteUrl };
            var watch = Stopwatch.StartNew();

            Log.Info(string.Format("Sync of '{0}' ({1}) to {2} started: {3} tasks, {4} resources, {5} assignments.",
                snapshot.ProjectName, snapshot.ProjectId, siteUrl, snapshot.Tasks.Count, snapshot.Resources.Count, snapshot.Assignments.Count));

            SiteInfo site;
            try
            {
                site = await _sites.ResolveAsync(siteUrl, ct).ConfigureAwait(false);
            }
            catch
            {
                _sites.Invalidate(siteUrl);
                throw;
            }

            var work = new List<KeyValuePair<ListSchema, IList<SyncRow>>>
            {
                new KeyValuePair<ListSchema, IList<SyncRow>>(ListSchemas.Projects(settings), new[] { snapshot.Project })
            };
            if (settings.SyncTasks) work.Add(new KeyValuePair<ListSchema, IList<SyncRow>>(ListSchemas.Tasks(settings), snapshot.Tasks));
            if (settings.SyncResources) work.Add(new KeyValuePair<ListSchema, IList<SyncRow>>(ListSchemas.Resources(settings), snapshot.Resources));
            if (settings.SyncAssignments) work.Add(new KeyValuePair<ListSchema, IList<SyncRow>>(ListSchemas.Assignments(settings), snapshot.Assignments));

            foreach (var item in work)
            {
                ct.ThrowIfCancellationRequested();
                try
                {
                    var listResult = await SyncListAsync(site.Id, item.Key, snapshot.ProjectId, item.Value, settings, ct).ConfigureAwait(false);
                    result.Lists.Add(listResult);
                }
                catch (OperationCanceledException)
                {
                    throw;
                }
                catch (Exception ex)
                {
                    // Force re-discovery of lists/columns on the next run.
                    _provisioner.InvalidateAll();
                    Log.Error("Synchronization of list '" + item.Key.Title + "' failed.", ex);
                    var failed = new ListSyncResult(item.Key.Title);
                    failed.Errors.Add(ex.Message);
                    result.Lists.Add(failed);
                }
            }

            watch.Stop();
            result.Duration = watch.Elapsed;
            Log.Info(result.Summary);
            return result;
        }

        private async Task<ListSyncResult> SyncListAsync(
            string siteId, ListSchema schema, string projectId, IList<SyncRow> rows, SyncSettings settings, CancellationToken ct)
        {
            var result = new ListSyncResult(schema.Title);
            var listId = await _provisioner.EnsureListAsync(siteId, schema, settings.AutoProvisionLists, ct).ConfigureAwait(false);
            var itemsUrl = "/sites/" + siteId + "/lists/" + listId + "/items";

            var existing = await LoadExistingAsync(itemsUrl, projectId, ct).ConfigureAwait(false);

            var requests = new List<BatchRequest>();
            var kinds = new List<char>();
            var seen = new HashSet<string>(StringComparer.Ordinal);

            foreach (var row in rows.Where(r => r != null))
            {
                if (!seen.Add(row.Key))
                {
                    Log.Warn("Duplicate key '" + row.Key + "' in list '" + schema.Title + "' skipped.");
                    continue;
                }

                var hash = RowHasher.Compute(row, schema.Version);
                List<ExistingItem> matches;
                if (!existing.TryGetValue(row.Key, out matches))
                {
                    requests.Add(new BatchRequest("POST", itemsUrl, new JObject { ["fields"] = ToFields(row, schema, projectId, hash, forCreate: true) }));
                    kinds.Add('A');
                    continue;
                }

                existing.Remove(row.Key);
                var first = matches[0];
                if (first.Hash == hash)
                {
                    result.Unchanged++;
                }
                else
                {
                    requests.Add(new BatchRequest("PATCH", itemsUrl + "/" + first.Id + "/fields", ToFields(row, schema, projectId, hash, forCreate: false)));
                    kinds.Add('U');
                }

                // Remove accidental duplicates (e.g. from an interrupted earlier run).
                foreach (var duplicate in matches.Skip(1))
                {
                    requests.Add(new BatchRequest("DELETE", itemsUrl + "/" + duplicate.Id));
                    kinds.Add('D');
                }
            }

            if (settings.DeleteRemovedItems)
            {
                foreach (var orphan in existing.Values.SelectMany(v => v))
                {
                    requests.Add(new BatchRequest("DELETE", itemsUrl + "/" + orphan.Id));
                    kinds.Add('D');
                }
            }

            if (requests.Count == 0) return result;

            var responses = await _graph.BatchAsync(requests, ct).ConfigureAwait(false);
            for (var i = 0; i < responses.Count; i++)
            {
                var response = responses[i];
                if (response != null && (response.IsSuccess || (kinds[i] == 'D' && response.Status == 404)))
                {
                    switch (kinds[i])
                    {
                        case 'A': result.Added++; break;
                        case 'U': result.Updated++; break;
                        default: result.Deleted++; break;
                    }
                }
                else
                {
                    var message = response == null ? "No response" : response.ErrorMessage;
                    result.Errors.Add(requests[i].Method + " " + requests[i].Url + " -> " + message);
                }
            }

            foreach (var error in result.Errors.Take(20))
            {
                Log.Warn(error);
            }

            return result;
        }

        private async Task<Dictionary<string, List<ExistingItem>>> LoadExistingAsync(string itemsUrl, string projectId, CancellationToken ct)
        {
            var filter = "fields/" + Columns.ProjectId + " eq '" + projectId.Replace("'", "''") + "'";
            var url = itemsUrl
                + "?$expand=fields($select=" + Columns.Key + "," + Columns.Hash + ")"
                + "&$filter=" + Uri.EscapeDataString(filter)
                + "&$top=999";

            var items = await _graph.GetAllAsync(url, ct).ConfigureAwait(false);
            var map = new Dictionary<string, List<ExistingItem>>(StringComparer.Ordinal);
            foreach (var item in items)
            {
                var fields = item["fields"] as JObject;
                var key = fields == null ? null : (string)fields[Columns.Key];
                var id = (string)item["id"];
                if (string.IsNullOrEmpty(id)) continue;

                // Items without a key cannot be matched; treat them as orphans of this project.
                key = string.IsNullOrEmpty(key) ? "\u0000orphan:" + id : key;

                List<ExistingItem> list;
                if (!map.TryGetValue(key, out list))
                {
                    list = new List<ExistingItem>();
                    map[key] = list;
                }

                list.Add(new ExistingItem(id, fields == null ? null : (string)fields[Columns.Hash]));
            }

            return map;
        }

        private static JObject ToFields(SyncRow row, ListSchema schema, string projectId, string hash, bool forCreate)
        {
            var fields = new JObject
            {
                ["Title"] = Truncate(row.Title, TextMaxLength),
                [Columns.ProjectId] = projectId,
                [Columns.Key] = row.Key,
                [Columns.Hash] = hash
            };

            foreach (var kv in row.Values)
            {
                var definition = schema.Find(kv.Key);
                if (definition == null) continue;

                var token = ToToken(kv.Value, definition);
                if (token == null)
                {
                    // On create, omit empty values; on update, clear them.
                    if (!forCreate) fields[kv.Key] = JValue.CreateNull();
                    continue;
                }

                fields[kv.Key] = token;
            }

            return fields;
        }

        private static JToken ToToken(object value, FieldDefinition definition)
        {
            if (value == null) return null;

            switch (definition.Kind)
            {
                case FieldKind.Text:
                    var text = Convert.ToString(value, CultureInfo.InvariantCulture);
                    return string.IsNullOrEmpty(text) ? null : new JValue(Truncate(text, TextMaxLength));

                case FieldKind.Note:
                    var note = Convert.ToString(value, CultureInfo.InvariantCulture);
                    return string.IsNullOrEmpty(note) ? null : new JValue(Truncate(note, NoteMaxLength));

                case FieldKind.Number:
                    double number;
                    if (value is double) number = (double)value;
                    else if (!double.TryParse(Convert.ToString(value, CultureInfo.InvariantCulture), NumberStyles.Float, CultureInfo.InvariantCulture, out number)) return null;
                    if (double.IsNaN(number) || double.IsInfinity(number)) return null;
                    return new JValue(Math.Round(number, 6));

                case FieldKind.DateTime:
                    if (!(value is DateTime)) return null;
                    var utc = ((DateTime)value).ToUniversalTime();
                    return new JValue(utc.ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", CultureInfo.InvariantCulture));

                case FieldKind.Boolean:
                    if (value is bool) return new JValue((bool)value);
                    return null;

                default:
                    return null;
            }
        }

        private static string Truncate(string value, int max)
        {
            if (value == null) return string.Empty;
            return value.Length <= max ? value : value.Substring(0, max);
        }

        private sealed class ExistingItem
        {
            public ExistingItem(string id, string hash)
            {
                Id = id;
                Hash = hash;
            }

            public string Id { get; private set; }

            public string Hash { get; private set; }
        }
    }
}
