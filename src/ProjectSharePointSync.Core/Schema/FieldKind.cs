namespace ProjectSharePointSync.Core.Schema
{
    public enum FieldKind
    {
        /// <summary>Single line of text (max 255 characters).</summary>
        Text,

        /// <summary>Multiple lines of plain text.</summary>
        Note,

        Number,

        DateTime,

        Boolean
    }
}
