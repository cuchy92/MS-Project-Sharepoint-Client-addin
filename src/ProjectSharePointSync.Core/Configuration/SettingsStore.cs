using System;
using System.IO;
using Newtonsoft.Json;
using ProjectSharePointSync.Core.Diagnostics;

namespace ProjectSharePointSync.Core.Configuration
{
    /// <summary>
    /// Loads and saves <see cref="SyncSettings"/> in %APPDATA%\ProjectSharePointSync\settings.json.
    /// </summary>
    public static class SettingsStore
    {
        public static string SettingsDirectory
        {
            get
            {
                return Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "ProjectSharePointSync");
            }
        }

        public static string SettingsFile
        {
            get { return Path.Combine(SettingsDirectory, "settings.json"); }
        }

        public static SyncSettings Load()
        {
            try
            {
                if (File.Exists(SettingsFile))
                {
                    var json = File.ReadAllText(SettingsFile);
                    var settings = JsonConvert.DeserializeObject<SyncSettings>(json);
                    if (settings != null)
                    {
                        return settings;
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Error("Failed to load settings, using defaults.", ex);
            }

            return new SyncSettings();
        }

        public static void Save(SyncSettings settings)
        {
            if (settings == null) throw new ArgumentNullException("settings");
            Directory.CreateDirectory(SettingsDirectory);
            var json = JsonConvert.SerializeObject(settings, Formatting.Indented);
            var tmp = SettingsFile + ".tmp";
            File.WriteAllText(tmp, json);
            if (File.Exists(SettingsFile))
            {
                File.Replace(tmp, SettingsFile, null);
            }
            else
            {
                File.Move(tmp, SettingsFile);
            }
        }
    }
}
