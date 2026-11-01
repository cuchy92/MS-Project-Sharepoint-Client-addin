namespace ProjectSharePointSync.Core.Schema
{
    /// <summary>
    /// Internal names of the SharePoint columns. All custom columns are prefixed with "Ps"
    /// to avoid collisions with built-in SharePoint columns.
    /// </summary>
    public static class Columns
    {
        // Common (technical) columns
        public const string ProjectId = "PsProjectId";
        public const string Key = "PsKey";
        public const string Hash = "PsHash";
        public const string ProjectName = "PsProjectName";

        // Shared business columns
        public const string UniqueId = "PsUniqueId";
        public const string Start = "PsStart";
        public const string Finish = "PsFinish";
        public const string DurationDays = "PsDurationDays";
        public const string WorkHours = "PsWorkHours";
        public const string ActualWorkHours = "PsActualWorkHours";
        public const string Cost = "PsCost";
        public const string PercentComplete = "PsPercentComplete";

        // Project
        public const string FilePath = "PsFilePath";
        public const string Manager = "PsManager";
        public const string Company = "PsCompany";
        public const string Author = "PsAuthor";
        public const string StatusDate = "PsStatusDate";
        public const string TaskCount = "PsTaskCount";
        public const string ResourceCount = "PsResourceCount";
        public const string LastSync = "PsLastSync";
        public const string LastSyncBy = "PsLastSyncBy";

        // Task
        public const string TaskId = "PsTaskId";
        public const string Wbs = "PsWbs";
        public const string OutlineLevel = "PsOutlineLevel";
        public const string ParentUniqueId = "PsParentUniqueId";
        public const string ActualStart = "PsActualStart";
        public const string ActualFinish = "PsActualFinish";
        public const string BaselineStart = "PsBaselineStart";
        public const string BaselineFinish = "PsBaselineFinish";
        public const string Deadline = "PsDeadline";
        public const string PercentWorkComplete = "PsPercentWorkComplete";
        public const string Summary = "PsSummary";
        public const string Milestone = "PsMilestone";
        public const string Critical = "PsCritical";
        public const string Active = "PsActive";
        public const string ResourceNames = "PsResourceNames";
        public const string Predecessors = "PsPredecessors";
        public const string Notes = "PsNotes";

        // Resource
        public const string ResourceId = "PsResourceId";
        public const string ResourceType = "PsResourceType";
        public const string Initials = "PsInitials";
        public const string Group = "PsGroup";
        public const string Email = "PsEmail";
        public const string MaxUnits = "PsMaxUnits";
        public const string StandardRate = "PsStandardRate";

        // Assignment
        public const string TaskUniqueId = "PsTaskUniqueId";
        public const string ResourceUniqueId = "PsResourceUniqueId";
        public const string TaskName = "PsTaskName";
        public const string ResourceName = "PsResourceName";
        public const string Units = "PsUnits";

        /// <summary>Prefix for user-configured custom task fields (e.g. PsCfText1).</summary>
        public const string CustomFieldPrefix = "PsCf";
    }
}
