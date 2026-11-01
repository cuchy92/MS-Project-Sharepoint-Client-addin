using System;
using System.Collections.Generic;

namespace ProjectSharePointSync.Core.Model
{
    /// <summary>
    /// Immutable-by-convention copy of the project data, read on the UI thread
    /// and then synchronized to SharePoint on a background thread
    /// (COM objects must not be touched from the background thread).
    /// </summary>
    public sealed class ProjectSnapshot
    {
        public ProjectSnapshot()
        {
            Tasks = new List<SyncRow>();
            Resources = new List<SyncRow>();
            Assignments = new List<SyncRow>();
            CapturedAt = DateTime.UtcNow;
        }

        /// <summary>Stable project identifier (stored as a custom document property in the .mpp).</summary>
        public string ProjectId { get; set; }

        public string ProjectName { get; set; }

        /// <summary>Optional per-project site URL (custom document property "PsSyncSiteUrl").</summary>
        public string SiteUrlOverride { get; set; }

        public DateTime CapturedAt { get; set; }

        public SyncRow Project { get; set; }

        public List<SyncRow> Tasks { get; private set; }

        public List<SyncRow> Resources { get; private set; }

        public List<SyncRow> Assignments { get; private set; }
    }
}
