using System.Collections.Generic;
using System.Linq;
using System.Text;
using ProjectSharePointSync.Core.Configuration;

namespace ProjectSharePointSync.Core.Schema
{
    /// <summary>
    /// Definitions of the four SharePoint lists the add-in writes to.
    /// </summary>
    public static class ListSchemas
    {
        private static IEnumerable<FieldDefinition> Common()
        {
            yield return new FieldDefinition(Columns.ProjectId, "Project ID", FieldKind.Text, indexed: true);
            yield return new FieldDefinition(Columns.Key, "Sync Key", FieldKind.Text, indexed: true);
            yield return new FieldDefinition(Columns.Hash, "Sync Hash", FieldKind.Text);
        }

        public static ListSchema Projects(SyncSettings settings)
        {
            var fields = Common().Concat(new[]
            {
                new FieldDefinition(Columns.FilePath, "File Path", FieldKind.Note),
                new FieldDefinition(Columns.Start, "Start", FieldKind.DateTime),
                new FieldDefinition(Columns.Finish, "Finish", FieldKind.DateTime),
                new FieldDefinition(Columns.DurationDays, "Duration (days)", FieldKind.Number),
                new FieldDefinition(Columns.WorkHours, "Work (h)", FieldKind.Number),
                new FieldDefinition(Columns.ActualWorkHours, "Actual Work (h)", FieldKind.Number),
                new FieldDefinition(Columns.Cost, "Cost", FieldKind.Number),
                new FieldDefinition(Columns.PercentComplete, "% Complete", FieldKind.Number),
                new FieldDefinition(Columns.Manager, "Manager", FieldKind.Text),
                new FieldDefinition(Columns.Company, "Company", FieldKind.Text),
                new FieldDefinition(Columns.Author, "Author", FieldKind.Text),
                new FieldDefinition(Columns.StatusDate, "Status Date", FieldKind.DateTime),
                new FieldDefinition(Columns.TaskCount, "Task Count", FieldKind.Number),
                new FieldDefinition(Columns.ResourceCount, "Resource Count", FieldKind.Number),
                new FieldDefinition(Columns.LastSync, "Last Sync", FieldKind.DateTime),
                new FieldDefinition(Columns.LastSyncBy, "Last Sync By", FieldKind.Text),
            });
            return new ListSchema(settings.ProjectsListTitle, "Projects synchronized from Microsoft Project desktop.", fields);
        }

        public static ListSchema Tasks(SyncSettings settings)
        {
            var fields = Common().Concat(new[]
            {
                new FieldDefinition(Columns.ProjectName, "Project", FieldKind.Text),
                new FieldDefinition(Columns.UniqueId, "Unique ID", FieldKind.Number),
                new FieldDefinition(Columns.TaskId, "ID", FieldKind.Number),
                new FieldDefinition(Columns.Wbs, "WBS", FieldKind.Text),
                new FieldDefinition(Columns.OutlineLevel, "Outline Level", FieldKind.Number),
                new FieldDefinition(Columns.ParentUniqueId, "Parent Unique ID", FieldKind.Number),
                new FieldDefinition(Columns.Start, "Start", FieldKind.DateTime),
                new FieldDefinition(Columns.Finish, "Finish", FieldKind.DateTime),
                new FieldDefinition(Columns.ActualStart, "Actual Start", FieldKind.DateTime),
                new FieldDefinition(Columns.ActualFinish, "Actual Finish", FieldKind.DateTime),
                new FieldDefinition(Columns.BaselineStart, "Baseline Start", FieldKind.DateTime),
                new FieldDefinition(Columns.BaselineFinish, "Baseline Finish", FieldKind.DateTime),
                new FieldDefinition(Columns.Deadline, "Deadline", FieldKind.DateTime),
                new FieldDefinition(Columns.DurationDays, "Duration (days)", FieldKind.Number),
                new FieldDefinition(Columns.WorkHours, "Work (h)", FieldKind.Number),
                new FieldDefinition(Columns.ActualWorkHours, "Actual Work (h)", FieldKind.Number),
                new FieldDefinition(Columns.Cost, "Cost", FieldKind.Number),
                new FieldDefinition(Columns.PercentComplete, "% Complete", FieldKind.Number),
                new FieldDefinition(Columns.PercentWorkComplete, "% Work Complete", FieldKind.Number),
                new FieldDefinition(Columns.Summary, "Summary", FieldKind.Boolean),
                new FieldDefinition(Columns.Milestone, "Milestone", FieldKind.Boolean),
                new FieldDefinition(Columns.Critical, "Critical", FieldKind.Boolean),
                new FieldDefinition(Columns.Active, "Active", FieldKind.Boolean),
                new FieldDefinition(Columns.ResourceNames, "Resource Names", FieldKind.Note),
                new FieldDefinition(Columns.Predecessors, "Predecessors", FieldKind.Text),
                new FieldDefinition(Columns.Notes, "Notes", FieldKind.Note),
            }).Concat(CustomTaskFieldDefinitions(settings));
            return new ListSchema(settings.TasksListTitle, "Tasks synchronized from Microsoft Project desktop.", fields);
        }

        public static ListSchema Resources(SyncSettings settings)
        {
            var fields = Common().Concat(new[]
            {
                new FieldDefinition(Columns.ProjectName, "Project", FieldKind.Text),
                new FieldDefinition(Columns.UniqueId, "Unique ID", FieldKind.Number),
                new FieldDefinition(Columns.ResourceId, "ID", FieldKind.Number),
                new FieldDefinition(Columns.ResourceType, "Type", FieldKind.Text),
                new FieldDefinition(Columns.Initials, "Initials", FieldKind.Text),
                new FieldDefinition(Columns.Group, "Group", FieldKind.Text),
                new FieldDefinition(Columns.Email, "E-mail", FieldKind.Text),
                new FieldDefinition(Columns.MaxUnits, "Max Units (%)", FieldKind.Number),
                new FieldDefinition(Columns.StandardRate, "Standard Rate", FieldKind.Text),
                new FieldDefinition(Columns.WorkHours, "Work (h)", FieldKind.Number),
                new FieldDefinition(Columns.ActualWorkHours, "Actual Work (h)", FieldKind.Number),
                new FieldDefinition(Columns.Cost, "Cost", FieldKind.Number),
            });
            return new ListSchema(settings.ResourcesListTitle, "Resources synchronized from Microsoft Project desktop.", fields);
        }

        public static ListSchema Assignments(SyncSettings settings)
        {
            var fields = Common().Concat(new[]
            {
                new FieldDefinition(Columns.ProjectName, "Project", FieldKind.Text),
                new FieldDefinition(Columns.UniqueId, "Unique ID", FieldKind.Number),
                new FieldDefinition(Columns.TaskUniqueId, "Task Unique ID", FieldKind.Number),
                new FieldDefinition(Columns.ResourceUniqueId, "Resource Unique ID", FieldKind.Number),
                new FieldDefinition(Columns.TaskName, "Task", FieldKind.Text),
                new FieldDefinition(Columns.ResourceName, "Resource", FieldKind.Text),
                new FieldDefinition(Columns.Start, "Start", FieldKind.DateTime),
                new FieldDefinition(Columns.Finish, "Finish", FieldKind.DateTime),
                new FieldDefinition(Columns.WorkHours, "Work (h)", FieldKind.Number),
                new FieldDefinition(Columns.ActualWorkHours, "Actual Work (h)", FieldKind.Number),
                new FieldDefinition(Columns.Units, "Units (%)", FieldKind.Number),
                new FieldDefinition(Columns.Cost, "Cost", FieldKind.Number),
                new FieldDefinition(Columns.PercentWorkComplete, "% Work Complete", FieldKind.Number),
            });
            return new ListSchema(settings.AssignmentsListTitle, "Assignments synchronized from Microsoft Project desktop.", fields);
        }

        /// <summary>Column internal name used for a configured custom task field.</summary>
        public static string CustomFieldColumnName(string projectFieldName)
        {
            var sb = new StringBuilder(Columns.CustomFieldPrefix);
            foreach (var ch in projectFieldName ?? string.Empty)
            {
                if ((ch >= 'a' && ch <= 'z') || (ch >= 'A' && ch <= 'Z') || (ch >= '0' && ch <= '9'))
                {
                    sb.Append(ch);
                }
            }

            // SharePoint internal names are limited to 32 characters.
            return sb.Length > 32 ? sb.ToString(0, 32) : sb.ToString();
        }

        private static IEnumerable<FieldDefinition> CustomTaskFieldDefinitions(SyncSettings settings)
        {
            var seen = new HashSet<string>();
            foreach (var name in settings.CustomTaskFields ?? new List<string>())
            {
                if (string.IsNullOrWhiteSpace(name)) continue;
                var column = CustomFieldColumnName(name.Trim());
                if (column.Length <= Columns.CustomFieldPrefix.Length || !seen.Add(column)) continue;
                yield return new FieldDefinition(column, name.Trim(), FieldKind.Note);
            }
        }
    }
}
