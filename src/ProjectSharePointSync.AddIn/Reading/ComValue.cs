using System;
using System.Globalization;

namespace ProjectSharePointSync.AddIn.Reading
{
    /// <summary>
    /// Conversions of COM VARIANT values returned by the MS Project object model.
    /// Many properties return object and may contain "NA" or empty strings instead of a value.
    /// </summary>
    internal static class ComValue
    {
        public static DateTime? ToDate(object value)
        {
            if (value == null || value is DBNull) return null;
            if (value is DateTime)
            {
                var date = (DateTime)value;
                // MS Project uses dates far in the future / 1984 sentinels for "NA" in some places.
                if (date.Year < 1984 || date.Year > 2149) return null;
                return DateTime.SpecifyKind(date, DateTimeKind.Local);
            }

            DateTime parsed;
            var text = value as string;
            if (!string.IsNullOrWhiteSpace(text) && DateTime.TryParse(text, CultureInfo.CurrentCulture, DateTimeStyles.AssumeLocal, out parsed))
            {
                return parsed;
            }

            return null;
        }

        public static double? ToDouble(object value)
        {
            if (value == null || value is DBNull) return null;
            if (value is double) return (double)value;
            if (value is float || value is int || value is short || value is long || value is decimal || value is byte)
            {
                return Convert.ToDouble(value, CultureInfo.InvariantCulture);
            }

            double parsed;
            var text = value as string;
            if (!string.IsNullOrWhiteSpace(text)
                && double.TryParse(text, NumberStyles.Any, CultureInfo.CurrentCulture, out parsed))
            {
                return parsed;
            }

            return null;
        }

        public static bool? ToBool(object value)
        {
            if (value == null || value is DBNull) return null;
            if (value is bool) return (bool)value;
            var number = ToDouble(value);
            if (number.HasValue) return number.Value != 0;
            bool parsed;
            if (bool.TryParse(Convert.ToString(value, CultureInfo.InvariantCulture), out parsed)) return parsed;
            return null;
        }

        public static string ToText(object value)
        {
            if (value == null || value is DBNull) return null;
            var text = Convert.ToString(value, CultureInfo.CurrentCulture);
            return string.IsNullOrWhiteSpace(text) ? null : text.Trim();
        }

        /// <summary>Minutes (MS Project internal unit) to hours.</summary>
        public static double? MinutesToHours(object value)
        {
            var minutes = ToDouble(value);
            return minutes.HasValue ? minutes.Value / 60.0 : (double?)null;
        }

        /// <summary>Minutes to working days using the project's hours-per-day setting.</summary>
        public static double? MinutesToDays(object value, double minutesPerDay)
        {
            var minutes = ToDouble(value);
            if (!minutes.HasValue || minutesPerDay <= 0) return null;
            return minutes.Value / minutesPerDay;
        }

        /// <summary>Fraction (1.0 = 100 %) to percent.</summary>
        public static double? FractionToPercent(object value)
        {
            var fraction = ToDouble(value);
            return fraction.HasValue ? fraction.Value * 100.0 : (double?)null;
        }

        /// <summary>Boxes nullable values as plain object (null stays null).</summary>
        public static object Box<T>(T? value) where T : struct
        {
            return value.HasValue ? (object)value.Value : null;
        }
    }
}
