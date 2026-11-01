using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ProjectSharePointSync.AddIn.Reading;
using ProjectSharePointSync.AddIn.UI;
using ProjectSharePointSync.Core.Diagnostics;
using Office = Microsoft.Office.Core;

namespace ProjectSharePointSync.AddIn.Ribbon
{
    /// <summary>
    /// Ribbon XML callbacks ("SharePoint" tab).
    /// </summary>
    [ComVisible(true)]
    public class SyncRibbon : Office.IRibbonExtensibility
    {
        private const string RibbonResource = "ProjectSharePointSync.AddIn.Ribbon.SyncRibbon.xml";

        private Office.IRibbonUI _ribbon;

        public string GetCustomUI(string ribbonId)
        {
            using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(RibbonResource))
            {
                if (stream == null) throw new InvalidOperationException("Ribbon resource not found: " + RibbonResource);
                using (var reader = new StreamReader(stream))
                {
                    return reader.ReadToEnd();
                }
            }
        }

        public void Ribbon_Load(Office.IRibbonUI ribbonUi)
        {
            _ribbon = ribbonUi;
        }

        private static ThisAddIn AddIn
        {
            get { return Globals.ThisAddIn; }
        }

        public void OnSyncNow(Office.IRibbonControl control)
        {
            Guard(() =>
            {
                var settings = AddIn.Coordinator.Settings;
                if (!settings.IsConfigured)
                {
                    MessageBox.Show(Win32Window.ProjectMainWindow,
                        "The SharePoint connection is not configured yet. Please fill in the site URL and the application (client) ID.",
                        "SharePoint", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    ShowSettings();
                    return;
                }

                var project = AddIn.ActiveProjectOrNull();
                if (project == null)
                {
                    MessageBox.Show(Win32Window.ProjectMainWindow, "No project is open.", "SharePoint", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                AddIn.Coordinator.SyncNow(project);
            });
        }

        public bool GetSyncOnSavePressed(Office.IRibbonControl control)
        {
            try
            {
                return AddIn.Coordinator.Settings.SyncOnSave;
            }
            catch
            {
                return false;
            }
        }

        public void OnToggleSyncOnSave(Office.IRibbonControl control, bool pressed)
        {
            Guard(() =>
            {
                var settings = AddIn.Coordinator.Settings;
                settings.SyncOnSave = pressed;
                AddIn.Coordinator.UpdateSettings(settings);
            });
        }

        public void OnSettings(Office.IRibbonControl control)
        {
            Guard(ShowSettings);
        }

        public void OnOpenSite(Office.IRibbonControl control)
        {
            Guard(() =>
            {
                var project = AddIn.ActiveProjectOrNull();
                var url = project == null ? null : ProjectIdentity.GetProperty(project, ProjectIdentity.SiteUrlProperty);
                if (string.IsNullOrWhiteSpace(url)) url = AddIn.Coordinator.Settings.SiteUrl;
                Uri uri;
                if (Uri.TryCreate(url, UriKind.Absolute, out uri) && uri.Scheme == Uri.UriSchemeHttps)
                {
                    Process.Start(new ProcessStartInfo(uri.AbsoluteUri) { UseShellExecute = true });
                }
            });
        }

        public void OnOpenLog(Office.IRibbonControl control)
        {
            Guard(() =>
            {
                Directory.CreateDirectory(Log.LogDirectory);
                Process.Start(new ProcessStartInfo(Log.LogDirectory) { UseShellExecute = true });
            });
        }

        private void ShowSettings()
        {
            var project = AddIn.ActiveProjectOrNull();
            var projectSiteUrl = project == null ? null : ProjectIdentity.GetProperty(project, ProjectIdentity.SiteUrlProperty);

            using (var form = new SettingsForm(AddIn.Coordinator, AddIn.Coordinator.Settings, project != null, projectSiteUrl))
            {
                if (form.ShowDialog(Win32Window.ProjectMainWindow) != DialogResult.OK) return;

                AddIn.Coordinator.UpdateSettings(form.Result);

                if (project != null && !string.Equals(projectSiteUrl ?? string.Empty, form.ProjectSiteUrl ?? string.Empty, StringComparison.Ordinal))
                {
                    ProjectIdentity.SetProperty(project, ProjectIdentity.SiteUrlProperty, form.ProjectSiteUrl ?? string.Empty);
                }
            }

            if (_ribbon != null) _ribbon.Invalidate();
        }

        private static void Guard(Action action)
        {
            try
            {
                action();
            }
            catch (Exception ex)
            {
                Log.Error("Ribbon action failed.", ex);
                MessageBox.Show(Win32Window.ProjectMainWindow, ex.Message, "SharePoint", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
