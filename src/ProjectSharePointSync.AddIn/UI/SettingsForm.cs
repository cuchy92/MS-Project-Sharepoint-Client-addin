using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using ProjectSharePointSync.Core.Configuration;
using ProjectSharePointSync.Core.Diagnostics;

namespace ProjectSharePointSync.AddIn.UI
{
    /// <summary>
    /// Settings dialog (built in code, no designer file needed).
    /// </summary>
    internal sealed class SettingsForm : Form
    {
        private readonly SyncCoordinator _coordinator;
        private readonly SyncSettings _original;

        private readonly TextBox _siteUrl = new TextBox();
        private readonly TextBox _tenantId = new TextBox();
        private readonly TextBox _clientId = new TextBox();
        private readonly TextBox _projectSiteUrl = new TextBox();
        private readonly TextBox _projectsList = new TextBox();
        private readonly TextBox _tasksList = new TextBox();
        private readonly TextBox _resourcesList = new TextBox();
        private readonly TextBox _assignmentsList = new TextBox();
        private readonly TextBox _customFields = new TextBox();

        private readonly CheckBox _syncOnSave = new CheckBox { Text = "Synchronize automatically when the project is saved" };
        private readonly CheckBox _syncTasks = new CheckBox { Text = "Tasks" };
        private readonly CheckBox _syncResources = new CheckBox { Text = "Resources" };
        private readonly CheckBox _syncAssignments = new CheckBox { Text = "Assignments" };
        private readonly CheckBox _autoProvision = new CheckBox { Text = "Create missing lists and columns automatically" };
        private readonly CheckBox _deleteRemoved = new CheckBox { Text = "Delete list items removed from the project" };
        private readonly CheckBox _showSuccess = new CheckBox { Text = "Show result in the status bar" };
        private readonly CheckBox _showErrors = new CheckBox { Text = "Show a dialog when synchronization fails" };

        private readonly Label _status = new Label { AutoSize = true, ForeColor = SystemColors.GrayText };

        public SettingsForm(SyncCoordinator coordinator, SyncSettings settings, bool hasActiveProject, string projectSiteUrl)
        {
            _coordinator = coordinator;
            _original = settings;

            Text = "SharePoint synchronization - settings";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            AutoScaleMode = AutoScaleMode.Dpi;
            AutoSize = true;
            AutoSizeMode = AutoSizeMode.GrowAndShrink;
            Padding = new Padding(12);

            var layout = new TableLayoutPanel
            {
                ColumnCount = 2,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Dock = DockStyle.Fill
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 460));

            AddHeader(layout, "Connection");
            AddRow(layout, "SharePoint site URL:", _siteUrl, "e.g. https://contoso.sharepoint.com/sites/pmo");
            AddRow(layout, "Tenant ID:", _tenantId, "Tenant GUID or domain (contoso.onmicrosoft.com); 'organizations' = any work account");
            AddRow(layout, "Application (client) ID:", _clientId, "Entra ID app registration - see README");

            _projectSiteUrl.Enabled = hasActiveProject;
            AddRow(layout, "Site URL for this project:", _projectSiteUrl,
                hasActiveProject ? "Optional - overrides the site above for the active project (stored in the .mpp)" : "Open a project to set a project-specific site");

            AddHeader(layout, "Synchronization");
            AddWide(layout, _syncOnSave);
            var what = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0) };
            what.Controls.AddRange(new Control[] { _syncTasks, _syncResources, _syncAssignments });
            foreach (Control c in what.Controls) c.AutoSize = true;
            AddRow(layout, "Synchronize:", what, null);
            AddWide(layout, _autoProvision);
            AddWide(layout, _deleteRemoved);
            AddWide(layout, _showSuccess);
            AddWide(layout, _showErrors);

            AddHeader(layout, "Lists");
            AddRow(layout, "Projects list:", _projectsList, null);
            AddRow(layout, "Tasks list:", _tasksList, null);
            AddRow(layout, "Resources list:", _resourcesList, null);
            AddRow(layout, "Assignments list:", _assignmentsList, null);

            _customFields.Multiline = true;
            _customFields.ScrollBars = ScrollBars.Vertical;
            _customFields.Height = 70;
            AddRow(layout, "Extra task fields:", _customFields, "One MS Project field per line, e.g. Text1, Number2, Flag1 or a custom field name");

            var buttons = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.RightToLeft, Dock = DockStyle.Fill, Margin = new Padding(0, 12, 0, 0) };
            var ok = new Button { Text = "OK", DialogResult = DialogResult.None, AutoSize = true };
            var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, AutoSize = true };
            var test = new Button { Text = "Test connection", AutoSize = true };
            var signOut = new Button { Text = "Sign out", AutoSize = true };
            buttons.Controls.AddRange(new Control[] { cancel, ok, signOut, test });
            layout.Controls.Add(_status, 0, layout.RowCount);
            layout.SetColumnSpan(_status, 2);
            layout.RowCount++;
            layout.Controls.Add(buttons, 0, layout.RowCount);
            layout.SetColumnSpan(buttons, 2);
            layout.RowCount++;

            Controls.Add(layout);
            AcceptButton = ok;
            CancelButton = cancel;

            ok.Click += OnOk;
            test.Click += OnTestConnection;
            signOut.Click += OnSignOut;

            LoadValues(settings, projectSiteUrl);
        }

        public SyncSettings Result { get; private set; }

        public string ProjectSiteUrl { get; private set; }

        private void LoadValues(SyncSettings s, string projectSiteUrl)
        {
            _siteUrl.Text = s.SiteUrl;
            _tenantId.Text = s.TenantId;
            _clientId.Text = s.ClientId;
            _projectSiteUrl.Text = projectSiteUrl ?? string.Empty;
            _syncOnSave.Checked = s.SyncOnSave;
            _syncTasks.Checked = s.SyncTasks;
            _syncResources.Checked = s.SyncResources;
            _syncAssignments.Checked = s.SyncAssignments;
            _autoProvision.Checked = s.AutoProvisionLists;
            _deleteRemoved.Checked = s.DeleteRemovedItems;
            _showSuccess.Checked = s.ShowSuccessNotification;
            _showErrors.Checked = s.ShowErrorDialog;
            _projectsList.Text = s.ProjectsListTitle;
            _tasksList.Text = s.TasksListTitle;
            _resourcesList.Text = s.ResourcesListTitle;
            _assignmentsList.Text = s.AssignmentsListTitle;
            _customFields.Text = string.Join(Environment.NewLine, s.CustomTaskFields ?? new System.Collections.Generic.List<string>());
        }

        private SyncSettings Collect()
        {
            var s = _original.Clone();
            s.SiteUrl = _siteUrl.Text.Trim().TrimEnd('/');
            s.TenantId = string.IsNullOrWhiteSpace(_tenantId.Text) ? "organizations" : _tenantId.Text.Trim();
            s.ClientId = _clientId.Text.Trim();
            s.SyncOnSave = _syncOnSave.Checked;
            s.SyncTasks = _syncTasks.Checked;
            s.SyncResources = _syncResources.Checked;
            s.SyncAssignments = _syncAssignments.Checked;
            s.AutoProvisionLists = _autoProvision.Checked;
            s.DeleteRemovedItems = _deleteRemoved.Checked;
            s.ShowSuccessNotification = _showSuccess.Checked;
            s.ShowErrorDialog = _showErrors.Checked;
            s.ProjectsListTitle = Required(_projectsList.Text, "MSP Projects");
            s.TasksListTitle = Required(_tasksList.Text, "MSP Tasks");
            s.ResourcesListTitle = Required(_resourcesList.Text, "MSP Resources");
            s.AssignmentsListTitle = Required(_assignmentsList.Text, "MSP Assignments");
            s.CustomTaskFields = _customFields.Lines
                .Select(l => l.Trim())
                .Where(l => l.Length > 0)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            return s;
        }

        private void OnOk(object sender, EventArgs e)
        {
            var s = Collect();
            var projectUrl = _projectSiteUrl.Text.Trim().TrimEnd('/');

            if (!string.IsNullOrEmpty(s.SiteUrl) && !IsHttpsUrl(s.SiteUrl))
            {
                MessageBox.Show(this, "The site URL must be an absolute https:// URL.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                _siteUrl.Focus();
                return;
            }

            if (projectUrl.Length > 0 && !IsHttpsUrl(projectUrl))
            {
                MessageBox.Show(this, "The project site URL must be an absolute https:// URL.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                _projectSiteUrl.Focus();
                return;
            }

            Result = s;
            ProjectSiteUrl = projectUrl.Length > 0 ? projectUrl : null;
            DialogResult = DialogResult.OK;
            Close();
        }

        private async void OnTestConnection(object sender, EventArgs e)
        {
            var s = Collect();
            var url = _projectSiteUrl.Text.Trim();
            if (url.Length > 0) s.SiteUrl = url;

            if (!s.IsConfigured)
            {
                _status.Text = "Fill in the site URL (https://...) and the application (client) ID first.";
                return;
            }

            var button = (Button)sender;
            button.Enabled = false;
            _status.Text = "Connecting... (a browser window may open for sign-in)";
            try
            {
                var site = await _coordinator.TestConnectionAsync(s);
                var user = await _coordinator.GetSignedInUserAsync(s);
                _status.Text = string.Format("OK - connected to '{0}' as {1}.", site.DisplayName, user ?? "?");
            }
            catch (Exception ex)
            {
                Log.Error("Connection test failed.", ex);
                _status.Text = "Failed: " + ex.Message;
            }
            finally
            {
                if (!button.IsDisposed) button.Enabled = true;
            }
        }

        private async void OnSignOut(object sender, EventArgs e)
        {
            var s = Collect();
            if (string.IsNullOrWhiteSpace(s.ClientId)) return;
            try
            {
                await _coordinator.SignOutAsync(s);
                _status.Text = "Signed out.";
            }
            catch (Exception ex)
            {
                _status.Text = "Sign-out failed: " + ex.Message;
            }
        }

        private static bool IsHttpsUrl(string value)
        {
            Uri uri;
            return Uri.TryCreate(value, UriKind.Absolute, out uri) && uri.Scheme == Uri.UriSchemeHttps;
        }

        private static string Required(string value, string fallback)
        {
            return string.IsNullOrWhiteSpace(value) ? fallback : value.Trim();
        }

        private static void AddHeader(TableLayoutPanel layout, string text)
        {
            var label = new Label
            {
                Text = text,
                AutoSize = true,
                Font = new Font(SystemFonts.MessageBoxFont, FontStyle.Bold),
                Margin = new Padding(0, layout.RowCount == 0 ? 0 : 12, 0, 4)
            };
            layout.Controls.Add(label, 0, layout.RowCount);
            layout.SetColumnSpan(label, 2);
            layout.RowCount++;
        }

        private static void AddRow(TableLayoutPanel layout, string caption, Control control, string hint)
        {
            var label = new Label { Text = caption, AutoSize = true, Anchor = AnchorStyles.Left | AnchorStyles.Top, Margin = new Padding(0, 6, 8, 0) };
            control.Dock = DockStyle.Fill;

            Control right = control;
            if (!string.IsNullOrEmpty(hint))
            {
                var panel = new TableLayoutPanel { ColumnCount = 1, AutoSize = true, Dock = DockStyle.Fill, Margin = new Padding(0) };
                panel.Controls.Add(control);
                panel.Controls.Add(new Label { Text = hint, AutoSize = true, ForeColor = SystemColors.GrayText, MaximumSize = new Size(460, 0) });
                right = panel;
            }

            layout.Controls.Add(label, 0, layout.RowCount);
            layout.Controls.Add(right, 1, layout.RowCount);
            layout.RowCount++;
        }

        private static void AddWide(TableLayoutPanel layout, Control control)
        {
            if (control is CheckBox) ((CheckBox)control).AutoSize = true;
            layout.Controls.Add(control, 1, layout.RowCount);
            layout.RowCount++;
        }
    }
}
