using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ProjectSharePointSync.Core.SharePoint
{
    public sealed class ListSyncResult
    {
        public ListSyncResult(string listTitle)
        {
            ListTitle = listTitle;
            Errors = new List<string>();
        }

        public string ListTitle { get; private set; }

        public int Added { get; set; }

        public int Updated { get; set; }

        public int Deleted { get; set; }

        public int Unchanged { get; set; }

        public List<string> Errors { get; private set; }

        public override string ToString()
        {
            return string.Format("{0}: +{1} ~{2} -{3} ={4}{5}", ListTitle, Added, Updated, Deleted, Unchanged,
                Errors.Count > 0 ? " (" + Errors.Count + " error(s))" : string.Empty);
        }
    }

    public sealed class SyncResult
    {
        public SyncResult(string projectName)
        {
            ProjectName = projectName;
            Lists = new List<ListSyncResult>();
        }

        public string ProjectName { get; private set; }

        public string SiteUrl { get; set; }

        public List<ListSyncResult> Lists { get; private set; }

        public TimeSpan Duration { get; set; }

        public bool HasErrors
        {
            get { return Lists.Any(l => l.Errors.Count > 0); }
        }

        public string Summary
        {
            get
            {
                var sb = new StringBuilder();
                sb.AppendFormat("'{0}' synchronized to SharePoint in {1:0.0}s. ", ProjectName, Duration.TotalSeconds);
                sb.Append(string.Join("; ", Lists.Select(l => l.ToString())));
                return sb.ToString();
            }
        }

        public string ErrorDetails
        {
            get
            {
                return string.Join(Environment.NewLine,
                    Lists.SelectMany(l => l.Errors.Take(10).Select(e => l.ListTitle + ": " + e)));
            }
        }
    }
}
