using System;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace ProjectSharePointSync.Core.Model
{
    /// <summary>
    /// Computes a content hash of a row. The hash is stored in a hidden column so that
    /// unchanged items are not re-sent to SharePoint on every save.
    /// </summary>
    public static class RowHasher
    {
        public static string Compute(SyncRow row, string schemaVersion)
        {
            var sb = new StringBuilder();
            sb.Append("v=").Append(schemaVersion).Append('\n');
            sb.Append("Title=").Append(row.Title).Append('\n');
            foreach (var kv in row.Values.OrderBy(k => k.Key, StringComparer.Ordinal))
            {
                sb.Append(kv.Key).Append('=').Append(Format(kv.Value)).Append('\n');
            }

            using (var sha = SHA256.Create())
            {
                var bytes = sha.ComputeHash(Encoding.UTF8.GetBytes(sb.ToString()));
                return Convert.ToBase64String(bytes);
            }
        }

        private static string Format(object value)
        {
            if (value == null) return "\u0000";
            if (value is DateTime) return ((DateTime)value).ToUniversalTime().ToString("o", CultureInfo.InvariantCulture);
            if (value is double) return ((double)value).ToString("R", CultureInfo.InvariantCulture);
            if (value is bool) return (bool)value ? "1" : "0";
            return Convert.ToString(value, CultureInfo.InvariantCulture);
        }
    }
}
