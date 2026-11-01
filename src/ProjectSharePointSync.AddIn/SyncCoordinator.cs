using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Identity.Client;
using ProjectSharePointSync.AddIn.Reading;
using ProjectSharePointSync.AddIn.UI;
using ProjectSharePointSync.Core.Auth;
using ProjectSharePointSync.Core.Configuration;
using ProjectSharePointSync.Core.Diagnostics;
using ProjectSharePointSync.Core.Model;
using ProjectSharePointSync.Core.SharePoint;
using MSProject = Microsoft.Office.Interop.MSProject;

namespace ProjectSharePointSync.AddIn
{
    /// <summary>
    /// Captures project snapshots on the UI thread and synchronizes them to SharePoint
    /// on a single background worker, so saving in MS Project is never blocked by the network.
    /// </summary>
    internal sealed class SyncCoordinator : IDisposable
    {
        private readonly MSProject.Application _app;
        private readonly UiDispatcher _ui;
        private readonly CancellationTokenSource _shutdown = new CancellationTokenSource();
        private readonly object _gate = new object();
        private readonly Dictionary<string, PendingSync> _pending = new Dictionary<string, PendingSync>(StringComparer.Ordinal);
        private readonly Dictionary<string, Connection> _connections = new Dictionary<string, Connection>(StringComparer.OrdinalIgnoreCase);
        private readonly System.Windows.Forms.Timer _statusReset;

        private SyncSettings _settings;
        private Task _worker = Task.FromResult(0);
        private bool _workerRunning;
        private volatile bool _interactiveDeclined;
        private volatile bool _shuttingDown;

        public SyncCoordinator(MSProject.Application app, SyncSettings settings)
        {
            _app = app;
            _settings = settings ?? new SyncSettings();
            _ui = new UiDispatcher();
            _statusReset = new System.Windows.Forms.Timer { Interval = 10000 };
            _statusReset.Tick += (s, e) =>
            {
                _statusReset.Stop();
                try { _app.StatusBar = false; } catch { }
            };
        }

        /// <summary>Copy of the current settings.</summary>
        public SyncSettings Settings
        {
            get { lock (_gate) return _settings.Clone(); }
        }

        public bool IsBusy
        {
            get { lock (_gate) return _workerRunning; }
        }

        public void UpdateSettings(SyncSettings settings)
        {
            SettingsStore.Save(settings);
            lock (_gate)
            {
                _settings = settings.Clone();
            }

            _interactiveDeclined = false;
            Log.Info("Settings updated.");
        }

        /// <summary>Called from the ProjectBeforeSave event (UI thread).</summary>
        public void OnBeforeSave(MSProject.Project project)
        {
            var settings = Settings;
            if (!settings.SyncOnSave) return;
            if (!settings.IsConfigured)
            {
                Log.Warn("Sync on save skipped: add-in is not configured.");
                return;
            }

            Capture(project, settings, manual: false);
        }

        /// <summary>Manual synchronization from the ribbon (UI thread).</summary>
        public void SyncNow(MSProject.Project project)
        {
            var settings = Settings;
            _interactiveDeclined = false;
            Capture(project, settings, manual: true);
        }

        public async Task<SiteInfo> TestConnectionAsync(SyncSettings settings)
        {
            var connection = GetConnection(settings);
            connection.AllowInteractive = true;
            return await connection.Service.TestConnectionAsync(settings.SiteUrl, _shutdown.Token).ConfigureAwait(false);
        }

        public async Task SignOutAsync(SyncSettings settings)
        {
            await GetConnection(settings).Tokens.SignOutAsync().ConfigureAwait(false);
            Log.Info("Signed out.");
        }

        public Task<string> GetSignedInUserAsync(SyncSettings settings)
        {
            return GetConnection(settings).Tokens.GetSignedInUserAsync();
        }

        private void Capture(MSProject.Project project, SyncSettings settings, bool manual)
        {
            if (project == null) return;
            var watch = Stopwatch.StartNew();
            ProjectSnapshot snapshot;
            try
            {
                SetStatus("SharePoint: reading project data...", false);
                snapshot = new ProjectReader(_app).Read(project, settings);
            }
            catch (Exception ex)
            {
                Log.Error("Reading project data failed.", ex);
                Notify("Reading the project failed: " + ex.Message, isError: true, settings: settings);
                return;
            }

            Log.Info(string.Format("Snapshot of '{0}' captured in {1} ms.", snapshot.ProjectName, watch.ElapsedMilliseconds));
            Enqueue(new PendingSync(snapshot, settings, manual));
        }

        private void Enqueue(PendingSync job)
        {
            lock (_gate)
            {
                PendingSync previous;
                if (_pending.TryGetValue(job.Snapshot.ProjectId, out previous) && previous.Manual)
                {
                    job.Manual = true;
                }

                // A newer snapshot of the same project replaces a not yet started one.
                _pending[job.Snapshot.ProjectId] = job;

                if (!_workerRunning)
                {
                    _workerRunning = true;
                    _worker = Task.Run(() => WorkerLoopAsync());
                }
            }

            SetStatus("SharePoint: synchronizing '" + job.Snapshot.ProjectName + "'...", false);
        }

        private async Task WorkerLoopAsync()
        {
            while (true)
            {
                PendingSync job;
                lock (_gate)
                {
                    if (_pending.Count == 0 || _shutdown.IsCancellationRequested)
                    {
                        _pending.Clear();
                        _workerRunning = false;
                        return;
                    }

                    job = _pending.Values.First();
                    _pending.Remove(job.Snapshot.ProjectId);
                }

                await RunAsync(job).ConfigureAwait(false);
            }
        }

        private async Task RunAsync(PendingSync job)
        {
            var settings = job.Settings;
            try
            {
                var connection = GetConnection(settings);
                connection.AllowInteractive = !_shuttingDown && (job.Manual || !_interactiveDeclined);

                var user = await connection.Tokens.GetSignedInUserAsync().ConfigureAwait(false);
                if (!string.IsNullOrEmpty(user) && job.Snapshot.Project != null)
                {
                    job.Snapshot.Project.Set(Core.Schema.Columns.LastSyncBy, user);
                }

                var result = await connection.Service.SyncAsync(job.Snapshot, settings, _shutdown.Token).ConfigureAwait(false);

                if (result.HasErrors)
                {
                    Notify(result.Summary + Environment.NewLine + Environment.NewLine + result.ErrorDetails, isError: true, settings: settings);
                }
                else
                {
                    Notify(result.Summary, isError: false, settings: settings);
                }
            }
            catch (OperationCanceledException)
            {
                Log.Warn("Synchronization cancelled.");
            }
            catch (MsalClientException ex) when (ex.ErrorCode == MsalError.AuthenticationCanceledError)
            {
                _interactiveDeclined = true;
                Log.Warn("Sign-in cancelled by the user; automatic sign-in prompts suspended until 'Sync now'.");
                Notify("SharePoint sign-in was cancelled. The project was NOT synchronized. Use 'Sync now' to try again.", isError: false, settings: settings);
            }
            catch (AuthenticationRequiredException ex)
            {
                Log.Warn(ex.Message);
                Notify(ex.Message, isError: false, settings: settings);
            }
            catch (Exception ex)
            {
                Log.Error("Synchronization failed.", ex);
                Notify("Synchronization to SharePoint failed: " + ex.Message, isError: true, settings: settings);
            }
        }

        private Connection GetConnection(SyncSettings settings)
        {
            var key = (settings.TenantId ?? string.Empty).Trim() + "|" + (settings.ClientId ?? string.Empty).Trim();
            lock (_gate)
            {
                Connection connection;
                if (!_connections.TryGetValue(key, out connection))
                {
                    connection = new Connection(settings.TenantId, settings.ClientId);
                    _connections[key] = connection;
                }

                connection.Tokens.ParentWindow = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle;
                return connection;
            }
        }

        private void Notify(string message, bool isError, SyncSettings settings)
        {
            if (isError) Log.Warn(message); else Log.Info(message);

            _ui.Post(() =>
            {
                var firstLine = message.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault() ?? message;
                SetStatusOnUiThread("SharePoint: " + firstLine, true);

                if (isError && settings.ShowErrorDialog && !_shuttingDown)
                {
                    MessageBox.Show(Win32Window.ProjectMainWindow, message + Environment.NewLine + Environment.NewLine + "Log: " + Log.LogDirectory,
                        "SharePoint synchronization", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            });
        }

        private void SetStatus(string text, bool autoReset)
        {
            // Called on the UI thread (Capture/Enqueue) - set directly.
            SetStatusOnUiThread(text, autoReset);
        }

        private void SetStatusOnUiThread(string text, bool autoReset)
        {
            try
            {
                var settings = Settings;
                if (!autoReset || settings.ShowSuccessNotification || text.IndexOf("fail", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    _app.StatusBar = text.Length > 250 ? text.Substring(0, 250) : text;
                }

                _statusReset.Stop();
                _statusReset.Start();
            }
            catch
            {
                // The status bar is not available in some states (e.g. while a cell is being edited).
            }
        }

        /// <summary>Gives a running synchronization a chance to finish when MS Project exits.</summary>
        public void Dispose()
        {
            _shuttingDown = true;
            Task worker;
            lock (_gate) worker = _worker;
            try
            {
                if (!worker.Wait(TimeSpan.FromSeconds(20)))
                {
                    Log.Warn("Synchronization still running at shutdown; cancelling.");
                    _shutdown.Cancel();
                    worker.Wait(TimeSpan.FromSeconds(3));
                }
            }
            catch (AggregateException)
            {
            }

            _statusReset.Dispose();
            _ui.Dispose();
            lock (_gate)
            {
                foreach (var connection in _connections.Values) connection.Graph.Dispose();
                _connections.Clear();
            }
        }

        private sealed class PendingSync
        {
            public PendingSync(ProjectSnapshot snapshot, SyncSettings settings, bool manual)
            {
                Snapshot = snapshot;
                Settings = settings;
                Manual = manual;
            }

            public ProjectSnapshot Snapshot { get; private set; }

            public SyncSettings Settings { get; private set; }

            public bool Manual { get; set; }
        }

        private sealed class Connection
        {
            public Connection(string tenantId, string clientId)
            {
                Tokens = new GraphTokenProvider(tenantId, clientId);
                Graph = new GraphClient(ct => Tokens.GetAccessTokenAsync(AllowInteractive, ct));
                Service = new ProjectSyncService(Graph);
            }

            public GraphTokenProvider Tokens { get; private set; }

            public GraphClient Graph { get; private set; }

            public ProjectSyncService Service { get; private set; }

            public volatile bool AllowInteractive;
        }
    }
}
