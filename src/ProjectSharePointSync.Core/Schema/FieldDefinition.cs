namespace ProjectSharePointSync.Core.Schema
{
    public sealed class FieldDefinition
    {
        public FieldDefinition(string internalName, string displayName, FieldKind kind, bool indexed = false, bool hidden = false)
        {
            InternalName = internalName;
            DisplayName = displayName;
            Kind = kind;
            Indexed = indexed;
            Hidden = hidden;
        }

        /// <summary>Column internal name (letters and digits only).</summary>
        public string InternalName { get; private set; }

        public string DisplayName { get; private set; }

        public FieldKind Kind { get; private set; }

        public bool Indexed { get; private set; }

        public bool Hidden { get; private set; }
    }
}
