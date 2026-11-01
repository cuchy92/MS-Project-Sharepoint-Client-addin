using System;
using Office = Microsoft.Office.Core;
using MSProject = Microsoft.Office.Interop.MSProject;

namespace ProjectSharePointSync.AddIn.Reading
{
    /// <summary>
    /// Stores a stable project identifier as a custom document property inside the .mpp file,
    /// so that renaming or moving the file keeps it linked to the same SharePoint items.
    /// </summary>
    internal static class ProjectIdentity
    {
        public const string ProjectIdProperty = "PsSyncProjectId";
        public const string SiteUrlProperty = "PsSyncSiteUrl";

        public static string GetOrCreateProjectId(MSProject.Project project)
        {
            var existing = GetProperty(project, ProjectIdProperty);
            Guid parsed;
            if (!string.IsNullOrWhiteSpace(existing) && Guid.TryParse(existing, out parsed))
            {
                return parsed.ToString("D");
            }

            var id = Guid.NewGuid().ToString("D");
            SetProperty(project, ProjectIdProperty, id);
            return id;
        }

        public static string GetProperty(MSProject.Project project, string name)
        {
            try
            {
                var props = (Office.DocumentProperties)project.CustomDocumentProperties;
                foreach (Office.DocumentProperty prop in props)
                {
                    if (string.Equals(prop.Name, name, StringComparison.OrdinalIgnoreCase))
                    {
                        return prop.Value == null ? null : prop.Value.ToString();
                    }
                }
            }
            catch
            {
                // Property collection not available (e.g. read-only/server project).
            }

            return null;
        }

        public static void SetProperty(MSProject.Project project, string name, string value)
        {
            var props = (Office.DocumentProperties)project.CustomDocumentProperties;
            foreach (Office.DocumentProperty prop in props)
            {
                if (string.Equals(prop.Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    prop.Value = value;
                    return;
                }
            }

            props.Add(name, false, Office.MsoDocProperties.msoPropertyTypeString, value);
        }
    }
}
