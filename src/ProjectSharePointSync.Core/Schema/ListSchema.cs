using System.Collections.Generic;
using System.Linq;

namespace ProjectSharePointSync.Core.Schema
{
    public sealed class ListSchema
    {
        public ListSchema(string title, string description, IEnumerable<FieldDefinition> fields)
        {
            Title = title;
            Description = description;
            Fields = fields.ToList().AsReadOnly();
            Version = ComputeVersion(Fields);
        }

        public string Title { get; private set; }

        public string Description { get; private set; }

        public IReadOnlyList<FieldDefinition> Fields { get; private set; }

        /// <summary>Changes whenever the set of columns changes; part of the row hash.</summary>
        public string Version { get; private set; }

        public FieldDefinition Find(string internalName)
        {
            return Fields.FirstOrDefault(f => f.InternalName == internalName);
        }

        private static string ComputeVersion(IEnumerable<FieldDefinition> fields)
        {
            return string.Join(",", fields.Select(f => f.InternalName + ":" + (int)f.Kind));
        }
    }
}
