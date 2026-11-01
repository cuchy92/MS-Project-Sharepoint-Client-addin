using System;
using System.Globalization;
using System.IO;
using System.Threading;

namespace ProjectSharePointSync.Core.Diagnostics
{
    /// <summary>
    /// Minimal thread-safe file logger. Logs go to %LOCALAPPDATA%\ProjectSharePointSync\logs.
    /// </summary>
    public static class Log
    {
        private static readonly object Gate = new object();

        public static string LogDirectory
        {
            get
            {
                return Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "ProjectSharePointSync", "logs");
            }
        }

        public static void Info(string message)
        {
            Write("INFO ", message, null);
        }

        public static void Warn(string message)
        {
            Write("WARN ", message, null);
        }

        public static void Error(string message, Exception ex = null)
        {
            Write("ERROR", message, ex);
        }

        private static void Write(string level, string message, Exception ex)
        {
            try
            {
                var now = DateTime.Now;
                var line = string.Format(
                    CultureInfo.InvariantCulture,
                    "{0:yyyy-MM-dd HH:mm:ss.fff} [{1}] [T{2}] {3}{4}{5}",
                    now, level, Thread.CurrentThread.ManagedThreadId, message,
                    ex == null ? string.Empty : Environment.NewLine,
                    ex == null ? string.Empty : ex.ToString());

                lock (Gate)
                {
                    Directory.CreateDirectory(LogDirectory);
                    var file = Path.Combine(LogDirectory, "sync-" + now.ToString("yyyyMMdd", CultureInfo.InvariantCulture) + ".log");
                    File.AppendAllText(file, line + Environment.NewLine);
                }
            }
            catch
            {
                // Logging must never break the host application.
            }
        }
    }
}
