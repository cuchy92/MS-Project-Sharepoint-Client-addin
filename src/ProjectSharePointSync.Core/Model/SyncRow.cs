using System;
using System.Collections.Generic;

namespace ProjectSharePointSync.Core.Model
{
    /// <summary>
    /// One record (= one SharePoint list item) produced from the project.
    /// Values are keyed by column internal name and may be
    /// <see cref="string"/>, <see cref="double"/>, <see cref="bool"/>, <see cref="DateTime"/> or null.
    /// </summary>
    public sealed class SyncRow
    {
        public SyncRow(string key, string title)
        {
            if (string.IsNullOrEmpty(key)) throw new ArgumentException("Key is required.", "key");
            Key = key;
            Title = title ?? string.Empty;
            Values = new Dictionary<string, object>(StringComparer.Ordinal);
        }

        /// <summary>Stable unique key of the record (stored in the PsKey column).</summary>
        public string Key { get; private set; }

        /// <summary>Value of the standard SharePoint Title column.</summary>
        public string Title { get; set; }

        public IDictionary<string, object> Values { get; private set; }

        public SyncRow Set(string column, object value)
        {
            Values[column] = value;
            return this;
        }
    }
}
