using System;
using ProjectSharePointSync.AddIn.Ribbon;
using ProjectSharePointSync.Core.Configuration;
using ProjectSharePointSync.Core.Diagnostics;
using MSProject = Microsoft.Office.Interop.MSProject;
using Office = Microsoft.Office.Core;

namespace ProjectSharePointSync.AddIn
{
    public partial class ThisAddIn
    {
        internal SyncCoordinator Coordinator { get; private set; }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            Log.Info("Add-in starting.");
            Coordinator = new SyncCoordinator(Application, SettingsStore.Load());
            Application.ProjectBeforeSave += Application_ProjectBeforeSave;
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            Log.Info("Add-in shutting down.");
            try
            {
                Application.ProjectBeforeSave -= Application_ProjectBeforeSave;
            }
            catch
            {
                // COM server may already be gone.
            }

            if (Coordinator != null)
            {
                Coordinator.Dispose();
            }
        }

        /// <summary>
        /// Fires before MS Project writes the file. The project data is captured here and
        /// uploaded in the background; the save itself is never cancelled or delayed by the network.
        /// </summary>
        private void Application_ProjectBeforeSave(MSProject.Project pj, bool SaveAsUi, ref bool Cancel)
        {
            try
            {
                Coordinator.OnBeforeSave(pj);
            }
            catch (Exception ex)
            {
                Log.Error("ProjectBeforeSave handler failed.", ex);
            }
        }

        internal MSProject.Project ActiveProjectOrNull()
        {
            try
            {
                return Application.Projects.Count > 0 ? Application.ActiveProject : null;
            }
            catch
            {
                return null;
            }
        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new SyncRibbon();
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
