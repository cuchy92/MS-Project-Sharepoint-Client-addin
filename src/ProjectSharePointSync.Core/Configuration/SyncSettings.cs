using System;
using System.Collections.Generic;
using Newtonsoft.Json;

namespace ProjectSharePointSync.Core.Configuration
{
    /// <summary>
    /// User settings of the add-in. Persisted as JSON by <see cref="SettingsStore"/>.
    /// </summary>
    public sealed class SyncSettings
    {
        public SyncSettings()
        {
            SiteUrl = string.Empty;
            TenantId = "organizations";
            ClientId = string.Empty;
            SyncOnSave = true;
            SyncTasks = true;
            SyncResources = true;
            SyncAssignments = true;
            AutoProvisionLists = true;
            DeleteRemovedItems = true;
            ShowSuccessNotification = true;
            ShowErrorDialog = true;
            ProjectsListTitle = "MSP Projects";
            TasksListTitle = "MSP Tasks";
            ResourcesListTitle = "MSP Resources";
            AssignmentsListTitle = "MSP Assignments";
            CustomTaskFields = new List<string>();
        }

        /// <summary>Absolute URL of the SharePoint Online site, e.g. https://contoso.sharepoint.com/sites/pmo</summary>
        public string SiteUrl { get; set; }

        /// <summary>Entra ID tenant (GUID, domain, or "organizations").</summary>
        public string TenantId { get; set; }

        /// <summary>Application (client) ID of the Entra ID app registration (public client).</summary>
        public string ClientId { get; set; }

        /// <summary>Synchronize automatically every time a project is saved.</summary>
        public bool SyncOnSave { get; set; }

        public bool SyncTasks { get; set; }

        public bool SyncResources { get; set; }

        public bool SyncAssignments { get; set; }

        /// <summary>Create missing lists and columns automatically (requires Sites.Manage.All).</summary>
        public bool AutoProvisionLists { get; set; }

        /// <summary>Delete list items whose task/resource/assignment no longer exists in the project.</summary>
        public bool DeleteRemovedItems { get; set; }

        public bool ShowSuccessNotification { get; set; }

        public bool ShowErrorDialog { get; set; }

        public string ProjectsListTitle { get; set; }

        public string TasksListTitle { get; set; }

        public string ResourcesListTitle { get; set; }

        public string AssignmentsListTitle { get; set; }

        /// <summary>
        /// Additional task fields to synchronize, by MS Project field name or alias
        /// (e.g. "Text1", "Number3", "Flag1", or a renamed custom field). Stored as text columns.
        /// </summary>
        public List<string> CustomTaskFields { get; set; }

        [JsonIgnore]
        public bool IsConfigured
        {
            get
            {
                Uri uri;
                return !string.IsNullOrWhiteSpace(ClientId)
                    && Uri.TryCreate(SiteUrl, UriKind.Absolute, out uri)
                    && uri.Scheme == Uri.UriSchemeHttps;
            }
        }

        public SyncSettings Clone()
        {
            var json = JsonConvert.SerializeObject(this);
            return JsonConvert.DeserializeObject<SyncSettings>(json);
        }
    }
}
