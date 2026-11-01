using System;
using System.Collections.Generic;
using System.Globalization;
using ProjectSharePointSync.Core.Configuration;
using ProjectSharePointSync.Core.Diagnostics;
using ProjectSharePointSync.Core.Model;
using ProjectSharePointSync.Core.Schema;
using MSProject = Microsoft.Office.Interop.MSProject;

namespace ProjectSharePointSync.AddIn.Reading
{
    /// <summary>
    /// Reads the open project through the MS Project COM object model into a <see cref="ProjectSnapshot"/>.
    /// Must be called on the Project UI (main) thread.
    /// </summary>
    internal sealed class ProjectReader
    {
        private readonly MSProject.Application _app;

        public ProjectReader(MSProject.Application app)
        {
            _app = app;
        }

        public ProjectSnapshot Read(MSProject.Project project, SyncSettings settings)
        {
            var snapshot = new ProjectSnapshot
            {
                ProjectId = ProjectIdentity.GetOrCreateProjectId(project),
                SiteUrlOverride = ProjectIdentity.GetProperty(project, ProjectIdentity.SiteUrlProperty),
                ProjectName = GetProjectName(project)
            };

            var minutesPerDay = (ComValue.ToDouble(project.HoursPerDay) ?? 8.0) * 60.0;
            var customFields = ResolveCustomFields(settings);

            var taskCount = 0;
            foreach (MSProject.Task task in project.Tasks)
            {
                // Blank rows in the Gantt chart are returned as null.
                if (task == null) continue;
                if (ComValue.ToBool(task.ExternalTask) == true) continue;
                taskCount++;

                if (settings.SyncTasks)
                {
                    snapshot.Tasks.Add(ReadTask(snapshot, task, minutesPerDay, customFields));
                }

                if (settings.SyncAssignments)
                {
                    foreach (MSProject.Assignment assignment in task.Assignments)
                    {
                        if (assignment == null) continue;
                        snapshot.Assignments.Add(ReadAssignment(snapshot, assignment));
                    }
                }
            }

            var resourceCount = 0;
            foreach (MSProject.Resource resource in project.Resources)
            {
                if (resource == null) continue;
                resourceCount++;
                if (settings.SyncResources)
                {
                    snapshot.Resources.Add(ReadResource(snapshot, resource));
                }
            }

            snapshot.Project = ReadProject(snapshot, project, minutesPerDay, taskCount, resourceCount);
            return snapshot;
        }

        private static string GetProjectName(MSProject.Project project)
        {
            string name = null;
            try
            {
                name = ComValue.ToText(project.ProjectSummaryTask.Name);
            }
            catch
            {
                // ignore
            }

            if (string.IsNullOrEmpty(name))
            {
                name = ComValue.ToText(project.Name) ?? "Project";
                if (name.EndsWith(".mpp", StringComparison.OrdinalIgnoreCase))
                {
                    name = name.Substring(0, name.Length - 4);
                }
            }

            return name;
        }

        private SyncRow ReadProject(ProjectSnapshot snapshot, MSProject.Project project, double minutesPerDay, int taskCount, int resourceCount)
        {
            var row = new SyncRow(snapshot.ProjectId, snapshot.ProjectName);
            var summary = project.ProjectSummaryTask;

            row.Set(Columns.FilePath, Safe(() => ComValue.ToText(project.FullName)))
               .Set(Columns.Start, ComValue.Box(ComValue.ToDate(project.ProjectStart)))
               .Set(Columns.Finish, ComValue.Box(ComValue.ToDate(project.ProjectFinish)))
               .Set(Columns.Manager, Safe(() => ComValue.ToText(project.Manager)))
               .Set(Columns.Company, Safe(() => ComValue.ToText(project.Company)))
               .Set(Columns.Author, Safe(() => ComValue.ToText(project.Author)))
               .Set(Columns.StatusDate, ComValue.Box(ComValue.ToDate(project.StatusDate)))
               .Set(Columns.TaskCount, (double)taskCount)
               .Set(Columns.ResourceCount, (double)resourceCount)
               .Set(Columns.LastSync, DateTime.UtcNow)
               .Set(Columns.LastSyncBy, Environment.UserName);

            if (summary != null)
            {
                row.Set(Columns.DurationDays, ComValue.Box(ComValue.MinutesToDays(summary.Duration, minutesPerDay)))
                   .Set(Columns.WorkHours, ComValue.Box(ComValue.MinutesToHours(summary.Work)))
                   .Set(Columns.ActualWorkHours, ComValue.Box(ComValue.MinutesToHours(summary.ActualWork)))
                   .Set(Columns.Cost, ComValue.Box(ComValue.ToDouble(summary.Cost)))
                   .Set(Columns.PercentComplete, ComValue.Box(ComValue.ToDouble(summary.PercentComplete)));
            }

            return row;
        }

        private static SyncRow ReadTask(ProjectSnapshot snapshot, MSProject.Task task, double minutesPerDay, IList<KeyValuePair<string, MSProject.PjField>> customFields)
        {
            var uniqueId = task.UniqueID;
            var row = new SyncRow(snapshot.ProjectId + ":T:" + Convert.ToString(uniqueId, CultureInfo.InvariantCulture), ComValue.ToText(task.Name) ?? string.Empty);

            int? parentUniqueId = null;
            try
            {
                var parent = task.OutlineParent;
                if (parent != null && parent.UniqueID != 0) parentUniqueId = parent.UniqueID;
            }
            catch
            {
                // top-level tasks
            }

            row.Set(Columns.ProjectName, snapshot.ProjectName)
               .Set(Columns.UniqueId, Convert.ToDouble(uniqueId, CultureInfo.InvariantCulture))
               .Set(Columns.TaskId, Convert.ToDouble(task.ID, CultureInfo.InvariantCulture))
               .Set(Columns.Wbs, ComValue.ToText(task.WBS))
               .Set(Columns.OutlineLevel, ComValue.Box(ComValue.ToDouble(task.OutlineLevel)))
               .Set(Columns.ParentUniqueId, parentUniqueId.HasValue ? (object)(double)parentUniqueId.Value : null)
               .Set(Columns.Start, ComValue.Box(ComValue.ToDate(task.Start)))
               .Set(Columns.Finish, ComValue.Box(ComValue.ToDate(task.Finish)))
               .Set(Columns.ActualStart, ComValue.Box(ComValue.ToDate(task.ActualStart)))
               .Set(Columns.ActualFinish, ComValue.Box(ComValue.ToDate(task.ActualFinish)))
               .Set(Columns.BaselineStart, ComValue.Box(ComValue.ToDate(task.BaselineStart)))
               .Set(Columns.BaselineFinish, ComValue.Box(ComValue.ToDate(task.BaselineFinish)))
               .Set(Columns.Deadline, ComValue.Box(ComValue.ToDate(task.Deadline)))
               .Set(Columns.DurationDays, ComValue.Box(ComValue.MinutesToDays(task.Duration, minutesPerDay)))
               .Set(Columns.WorkHours, ComValue.Box(ComValue.MinutesToHours(task.Work)))
               .Set(Columns.ActualWorkHours, ComValue.Box(ComValue.MinutesToHours(task.ActualWork)))
               .Set(Columns.Cost, ComValue.Box(ComValue.ToDouble(task.Cost)))
               .Set(Columns.PercentComplete, ComValue.Box(ComValue.ToDouble(task.PercentComplete)))
               .Set(Columns.PercentWorkComplete, ComValue.Box(ComValue.ToDouble(task.PercentWorkComplete)))
               .Set(Columns.Summary, ComValue.Box(ComValue.ToBool(task.Summary)))
               .Set(Columns.Milestone, ComValue.Box(ComValue.ToBool(task.Milestone)))
               .Set(Columns.Critical, ComValue.Box(ComValue.ToBool(task.Critical)))
               .Set(Columns.Active, ComValue.Box(ComValue.ToBool(task.Active)))
               .Set(Columns.ResourceNames, ComValue.ToText(task.ResourceNames))
               .Set(Columns.Predecessors, ComValue.ToText(task.Predecessors))
               .Set(Columns.Notes, ComValue.ToText(task.Notes));

            foreach (var custom in customFields)
            {
                string value = null;
                try
                {
                    value = ComValue.ToText(task.GetField(custom.Value));
                }
                catch (Exception ex)
                {
                    Log.Warn("Cannot read custom field '" + custom.Key + "': " + ex.Message);
                }

                row.Set(custom.Key, value);
            }

            return row;
        }

        private static SyncRow ReadResource(ProjectSnapshot snapshot, MSProject.Resource resource)
        {
            var uniqueId = resource.UniqueID;
            var row = new SyncRow(snapshot.ProjectId + ":R:" + Convert.ToString(uniqueId, CultureInfo.InvariantCulture), ComValue.ToText(resource.Name) ?? string.Empty);

            row.Set(Columns.ProjectName, snapshot.ProjectName)
               .Set(Columns.UniqueId, Convert.ToDouble(uniqueId, CultureInfo.InvariantCulture))
               .Set(Columns.ResourceId, Convert.ToDouble(resource.ID, CultureInfo.InvariantCulture))
               .Set(Columns.ResourceType, ResourceTypeName(resource.Type))
               .Set(Columns.Initials, ComValue.ToText(resource.Initials))
               .Set(Columns.Group, ComValue.ToText(resource.Group))
               .Set(Columns.Email, ComValue.ToText(resource.EMailAddress))
               .Set(Columns.MaxUnits, ComValue.Box(ComValue.FractionToPercent(resource.MaxUnits)))
               .Set(Columns.StandardRate, ComValue.ToText(resource.StandardRate))
               .Set(Columns.WorkHours, ComValue.Box(ComValue.MinutesToHours(resource.Work)))
               .Set(Columns.ActualWorkHours, ComValue.Box(ComValue.MinutesToHours(resource.ActualWork)))
               .Set(Columns.Cost, ComValue.Box(ComValue.ToDouble(resource.Cost)));

            return row;
        }

        private static SyncRow ReadAssignment(ProjectSnapshot snapshot, MSProject.Assignment assignment)
        {
            var uniqueId = assignment.UniqueID;
            var taskName = ComValue.ToText(assignment.TaskName) ?? string.Empty;
            var resourceName = ComValue.ToText(assignment.ResourceName) ?? string.Empty;
            var row = new SyncRow(
                snapshot.ProjectId + ":A:" + Convert.ToString(uniqueId, CultureInfo.InvariantCulture),
                resourceName.Length > 0 ? taskName + " - " + resourceName : taskName);

            row.Set(Columns.ProjectName, snapshot.ProjectName)
               .Set(Columns.UniqueId, Convert.ToDouble(uniqueId, CultureInfo.InvariantCulture))
               .Set(Columns.TaskUniqueId, Convert.ToDouble(assignment.TaskUniqueID, CultureInfo.InvariantCulture))
               .Set(Columns.ResourceUniqueId, Convert.ToDouble(assignment.ResourceUniqueID, CultureInfo.InvariantCulture))
               .Set(Columns.TaskName, taskName)
               .Set(Columns.ResourceName, resourceName)
               .Set(Columns.Start, ComValue.Box(ComValue.ToDate(assignment.Start)))
               .Set(Columns.Finish, ComValue.Box(ComValue.ToDate(assignment.Finish)))
               .Set(Columns.WorkHours, ComValue.Box(ComValue.MinutesToHours(assignment.Work)))
               .Set(Columns.ActualWorkHours, ComValue.Box(ComValue.MinutesToHours(assignment.ActualWork)))
               .Set(Columns.Units, ComValue.Box(ComValue.FractionToPercent(assignment.Units)))
               .Set(Columns.Cost, ComValue.Box(ComValue.ToDouble(assignment.Cost)))
               .Set(Columns.PercentWorkComplete, ComValue.Box(ComValue.ToDouble(assignment.PercentWorkComplete)));

            return row;
        }

        private IList<KeyValuePair<string, MSProject.PjField>> ResolveCustomFields(SyncSettings settings)
        {
            var result = new List<KeyValuePair<string, MSProject.PjField>>();
            if (!settings.SyncTasks || settings.CustomTaskFields == null) return result;

            foreach (var name in settings.CustomTaskFields)
            {
                if (string.IsNullOrWhiteSpace(name)) continue;
                try
                {
                    var field = (MSProject.PjField)_app.FieldNameToFieldConstant(name.Trim(), MSProject.PjFieldType.pjTask);
                    result.Add(new KeyValuePair<string, MSProject.PjField>(ListSchemas.CustomFieldColumnName(name.Trim()), field));
                }
                catch (Exception ex)
                {
                    Log.Warn("Unknown MS Project task field '" + name + "' ignored: " + ex.Message);
                }
            }

            return result;
        }

        private static string ResourceTypeName(MSProject.PjResourceTypes type)
        {
            switch (type)
            {
                case MSProject.PjResourceTypes.pjResourceTypeWork: return "Work";
                case MSProject.PjResourceTypes.pjResourceTypeMaterial: return "Material";
                case MSProject.PjResourceTypes.pjResourceTypeCost: return "Cost";
                default: return type.ToString();
            }
        }

        private static string Safe(Func<string> read)
        {
            try
            {
                return read();
            }
            catch
            {
                return null;
            }
        }
    }
}
