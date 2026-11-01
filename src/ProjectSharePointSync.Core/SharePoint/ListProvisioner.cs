using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using ProjectSharePointSync.Core.Diagnostics;
using ProjectSharePointSync.Core.Schema;

namespace ProjectSharePointSync.Core.SharePoint
{
    /// <summary>
    /// Finds the SharePoint lists and makes sure they contain all required columns,
    /// creating them when auto-provisioning is enabled.
    /// </summary>
    public sealed class ListProvisioner
    {
        private readonly GraphClient _graph;

        // siteId|listTitle|schemaVersion -> listId (verified in this session)
        private readonly ConcurrentDictionary<string, string> _verified = new ConcurrentDictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        public ListProvisioner(GraphClient graph)
        {
            _graph = graph;
        }

        public async Task<string> EnsureListAsync(string siteId, ListSchema schema, bool autoProvision, CancellationToken ct)
        {
            var cacheKey = siteId + "|" + schema.Title + "|" + schema.Version;
            string listId;
            if (_verified.TryGetValue(cacheKey, out listId)) return listId;

            listId = await FindListAsync(siteId, schema.Title, ct).ConfigureAwait(false);
            if (listId == null)
            {
                if (!autoProvision)
                {
                    throw new InvalidOperationException(string.Format(
                        "SharePoint list '{0}' does not exist and automatic list creation is disabled.", schema.Title));
                }

                listId = await CreateListAsync(siteId, schema, ct).ConfigureAwait(false);
            }
            else
            {
                await EnsureColumnsAsync(siteId, listId, schema, autoProvision, ct).ConfigureAwait(false);
            }

            _verified[cacheKey] = listId;
            return listId;
        }

        public void InvalidateAll()
        {
            _verified.Clear();
        }

        private async Task<string> FindListAsync(string siteId, string title, CancellationToken ct)
        {
            var lists = await _graph.GetAllAsync("/sites/" + siteId + "/lists?$select=id,displayName,name&$top=200", ct).ConfigureAwait(false);
            var match = lists.FirstOrDefault(l => string.Equals((string)l["displayName"], title, StringComparison.OrdinalIgnoreCase))
                     ?? lists.FirstOrDefault(l => string.Equals((string)l["name"], title, StringComparison.OrdinalIgnoreCase));
            return match == null ? null : (string)match["id"];
        }

        private async Task<string> CreateListAsync(string siteId, ListSchema schema, CancellationToken ct)
        {
            Log.Info("Creating SharePoint list '" + schema.Title + "'.");
            var body = new JObject
            {
                ["displayName"] = schema.Title,
                ["description"] = schema.Description,
                ["list"] = new JObject { ["template"] = "genericList" },
                ["columns"] = new JArray(schema.Fields.Select(ToColumnDefinition))
            };

            var created = await _graph.PostAsync("/sites/" + siteId + "/lists", body, ct).ConfigureAwait(false);
            var listId = (string)created["id"];

            // Verify: some tenants ignore column flags (e.g. indexed) on list creation.
            await EnsureColumnsAsync(siteId, listId, schema, true, ct).ConfigureAwait(false);
            return listId;
        }

        private async Task EnsureColumnsAsync(string siteId, string listId, ListSchema schema, bool autoProvision, CancellationToken ct)
        {
            var columnsUrl = "/sites/" + siteId + "/lists/" + listId + "/columns";
            var existing = await _graph.GetAllAsync(columnsUrl + "?$select=id,name,indexed", ct).ConfigureAwait(false);
            var byName = new Dictionary<string, JObject>(StringComparer.OrdinalIgnoreCase);
            foreach (var column in existing)
            {
                var name = (string)column["name"];
                if (name != null && !byName.ContainsKey(name)) byName[name] = column;
            }

            var missing = schema.Fields.Where(f => !byName.ContainsKey(f.InternalName)).ToList();
            if (missing.Count > 0 && !autoProvision)
            {
                throw new InvalidOperationException(string.Format(
                    "SharePoint list '{0}' is missing columns: {1}. Enable automatic list creation or add them manually.",
                    schema.Title, string.Join(", ", missing.Select(m => m.InternalName))));
            }

            foreach (var field in missing)
            {
                Log.Info(string.Format("Adding column '{0}' to list '{1}'.", field.InternalName, schema.Title));
                await _graph.PostAsync(columnsUrl, ToColumnDefinition(field), ct).ConfigureAwait(false);
            }

            foreach (var field in schema.Fields.Where(f => f.Indexed))
            {
                JObject column;
                if (byName.TryGetValue(field.InternalName, out column) && !((bool?)column["indexed"] ?? false))
                {
                    if (!autoProvision)
                    {
                        Log.Warn(string.Format("Column '{0}' in list '{1}' is not indexed; large lists may fail to query.", field.InternalName, schema.Title));
                        continue;
                    }

                    Log.Info(string.Format("Indexing column '{0}' in list '{1}'.", field.InternalName, schema.Title));
                    await _graph.SendAsync(new System.Net.Http.HttpMethod("PATCH"), columnsUrl + "/" + (string)column["id"],
                        new JObject { ["indexed"] = true }, ct).ConfigureAwait(false);
                }
            }
        }

        private static JObject ToColumnDefinition(FieldDefinition field)
        {
            var column = new JObject
            {
                ["name"] = field.InternalName,
                ["displayName"] = field.DisplayName,
                ["enforceUniqueValues"] = false,
                ["hidden"] = field.Hidden,
                ["indexed"] = field.Indexed
            };

            switch (field.Kind)
            {
                case FieldKind.Text:
                    column["text"] = new JObject { ["maxLength"] = 255 };
                    break;
                case FieldKind.Note:
                    column["text"] = new JObject
                    {
                        ["allowMultipleLines"] = true,
                        ["textType"] = "plain",
                        ["linesForEditing"] = 6
                    };
                    break;
                case FieldKind.Number:
                    column["number"] = new JObject { ["decimalPlaces"] = "automatic" };
                    break;
                case FieldKind.DateTime:
                    column["dateTime"] = new JObject { ["format"] = "dateTime", ["displayAs"] = "standard" };
                    break;
                case FieldKind.Boolean:
                    column["boolean"] = new JObject();
                    break;
                default:
                    throw new ArgumentOutOfRangeException("field", field.Kind, "Unsupported field kind.");
            }

            return column;
        }
    }
}
