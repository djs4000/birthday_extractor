using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
namespace BirthdayExtractor
{
    /// <summary>
    /// Primary UI surface for configuring and running birthday extraction jobs.
    /// Hosts file selectors, date pickers, progress feedback, and history entry points.
    /// </summary>
    public class MainForm : Form
    {
        // --- UI controls ---
        private Panel content = null!;
        private TextBox txtCsv = null!;
        private Button btnBrowseCsv = null!;
        private RadioButton rbSourceCsv = null!;
        private RadioButton rbSourceOnline = null!;
        private DateTimePicker dtStart = null!;
        private DateTimePicker dtEnd = null!;
        private CheckBox chkCsv = null!;
        private CheckBox chkXlsx = null!;
        private TextBox txtOutDir = null!;
        private Button btnBrowseOut = null!;
        private Button btnRun = null!;
        private Button btnCancel = null!;
        private Button btnUpload = null!;
        private ProgressBar progress = null!;
        private RichTextBox txtLog = null!;
        private Label lblEnd = null!;
        // --- Processing state & config ---
        private readonly Processing _proc = new();
        private System.Threading.CancellationTokenSource? _cts;
        private AppConfig _cfg = null!;
        private ProcResult? _lastResult;
        private MenuStrip menu = null!;
        private ToolStripMenuItem miSettings = null!;
        private ToolStripMenuItem miHistory  = null!;
        /// <summary>
        /// Creates the main window, loads persisted configuration, and wires up all controls/events.
        /// </summary>
        public MainForm()
        {
            // 1) Load config FIRST, with a safe fallback so layout decisions can use persisted defaults
            try
            {
                _cfg = ConfigStore.LoadOrCreate() ?? new AppConfig();
            }
            catch (Exception ex)
            {
                _cfg = new AppConfig();
                LogRouter.LogException(ex, "Failed to load configuration");
            }
            // sanity defaults if someone hand-edited config
            if (_cfg.DefaultWindowDays <= 0) _cfg.DefaultWindowDays = 7;
            LogRouter.SetVerboseLoggingEnabled(_cfg.VerboseLoggingEnabled);
            // 2) Form shell: establish window chrome before wiring controls
            Text = $"Birthday Extractor v{AppVersion.Display}";
            Width = 820; Height = 600;
            MinimumSize = new Size(820, 600);
            StartPosition = FormStartPosition.CenterScreen;
            TryApplyWindowIcon();

            // 3) Build UI
            InitializeMenu();
            InitializeContentPanel();
            LogRouter.RegisterUiLogger(Log);

            // 4) Defaults pulled from config to pre-populate the form
            dtStart.Value = DateTime.Today.AddDays(_cfg.DefaultStartOffsetDays);
            dtEnd.Value   = dtStart.Value.AddDays(_cfg.DefaultWindowDays - 1);
            chkCsv.Checked  = _cfg.DefaultWriteCsv;
            chkXlsx.Checked = _cfg.DefaultWriteXlsx;
            dtStart.ValueChanged += (s, e) => dtEnd.Value = dtStart.Value.Date.AddDays(_cfg.DefaultWindowDays - 1);
            txtCsv.TextChanged += (s, e) => SyncDefaultOutDir();
            SyncDefaultOutDir();
            UpdateSourceUiState();
        }

        /// <summary>
        /// Sets up the main menu.
        /// </summary>
        private void InitializeMenu()
        {
            // 3) Menu (Dock Top) for settings + history shortcuts
            menu = new MenuStrip();
            miSettings = new ToolStripMenuItem("Settings...");
            miHistory  = new ToolStripMenuItem("View Processed History");
            miSettings.Click += (s, e) => OpenSettings();
            miHistory.Click  += (s, e) => ShowHistory();
            menu.Items.Add(miSettings);
            menu.Items.Add(miHistory);
            menu.Dock = DockStyle.Top;
            MainMenuStrip = menu;
            Controls.Add(menu);
        }

        /// <summary>
        /// Sets up the main content panel and all its child controls.
        /// </summary>
        private void InitializeContentPanel()
        {
            // 4) Content panel (Dock Fill) – all inputs go here
            content = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                Padding = new Padding(10)
            };
            Controls.Add(content);
            var menuHeight = menu?.GetPreferredSize(Size.Empty).Height
                             ?? menu?.Height
                             ?? 24;
            int topOffset = menuHeight + 6;
            int Offset(int baseTop) => topOffset + baseTop;
            // 5) Build your controls (labels, inputs, and action buttons)
            var lblSource = new Label { Left = 20, Top = Offset(12), Width = 120, Text = "Data Source:" };
            rbSourceCsv = new RadioButton { Left = 140, Top = Offset(10), Width = 150, Text = "Load from CSV", Checked = true };
            rbSourceOnline = new RadioButton { Left = 300, Top = Offset(10), Width = 220, Text = "Load from online source" };
            rbSourceCsv.CheckedChanged += (s, e) => UpdateSourceUiState();
            rbSourceOnline.CheckedChanged += (s, e) => UpdateSourceUiState();
            var lblCsv = new Label { Left = 20, Top = Offset(40), Width = 120, Text = "CSV File:" };
            txtCsv = new TextBox { Left = 140, Top = Offset(36), Width = 540, Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right };
            btnBrowseCsv = new Button { Left = 690, Top = Offset(34), Width = 90, Text = "Browse...", Anchor = AnchorStyles.Top | AnchorStyles.Right };
            btnBrowseCsv.Click += (s, e) => BrowseCsv();
            var lblStart = new Label { Left = 20, Top = Offset(70), Width = 120, Text = "Start Date:" };
            dtStart = new DateTimePicker { Left = 140, Top = Offset(66), Width = 200, Format = DateTimePickerFormat.Custom, CustomFormat = "yyyy-MM-dd" };
            dtEnd = new DateTimePicker { Left = 440, Top = Offset(66), Width = 200, Format = DateTimePickerFormat.Custom, CustomFormat = "yyyy-MM-dd" };
            lblEnd = new Label { Left = 360, Top = Offset(70), Width = 60, Text = "End Date:" };
            chkCsv = new CheckBox { Left = 140, Top = Offset(96), Width = 80, Text = "CSV" };
            chkXlsx = new CheckBox { Left = 230, Top = Offset(96), Width = 80, Text = "XLSX" };
            var lblOut = new Label { Left = 20, Top = Offset(136), Width = 120, Text = "Output Folder:" };
            txtOutDir = new TextBox { Left = 140, Top = Offset(132), Width = 540, Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right };
            btnBrowseOut = new Button { Left = 690, Top = Offset(130), Width = 90, Text = "Browse...", Anchor = AnchorStyles.Top | AnchorStyles.Right };
            btnBrowseOut.Click += (s, e) => BrowseOutDir();
            btnRun = new Button { Left = 140, Top = Offset(172), Width = 120, Text = "Run" };
            btnCancel = new Button { Left = 270, Top = Offset(172), Width = 120, Text = "Cancel", Enabled = false };
            btnUpload = new Button { Left = 400, Top = Offset(172), Width = 150, Text = "Upload to ERPNext", Enabled = false };
            btnRun.Click += async (s, e) => await RunAsync();
            btnCancel.Click += (s, e) => _cts?.Cancel();
            btnUpload.Click += async (s, e) => await UploadAsync();
            progress = new ProgressBar
            {
                Left = 20,
                Top = Offset(212),
                Width = 760,
                Height = 18,
                Style = ProgressBarStyle.Continuous,
                Minimum = 0,
                Maximum = 100,
                Value = 0,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            txtLog = new RichTextBox
            {
                Left = 20,
                Top = Offset(242),
                Width = 760,
                Height = 260,
                Multiline = true,
                ScrollBars = RichTextBoxScrollBars.Vertical,
                ReadOnly = true,
                DetectUrls = true,
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
            };
            txtLog.LinkClicked += (s, e) => OpenLogLink(e.LinkText);
            // 6) Add to content panel (not the form)
            content.Controls.AddRange(new Control[] {
                lblSource, rbSourceCsv, rbSourceOnline,
                lblCsv, txtCsv, btnBrowseCsv,
                lblStart, dtStart,
                lblEnd, dtEnd,
                chkCsv, chkXlsx,
                lblOut, txtOutDir, btnBrowseOut,
                btnRun, btnCancel, btnUpload,
                progress, txtLog
            });
        }

        /// <inheritdoc />
        protected override async void OnShown(EventArgs e)
        {
            base.OnShown(e);
            await CheckForUpdatesAsync();
        }
        /// <summary>
        /// Opens the secondary settings dialog and reapplies any updated defaults.
        /// </summary>
        private void OpenSettings()
        {
            using var dlg = new SettingsForm(_cfg);
            if (dlg.ShowDialog(this) == DialogResult.OK)
            {
                // Re-apply defaults after save
                dtStart.Value = DateTime.Today.AddDays(_cfg.DefaultStartOffsetDays);
                dtEnd.Value   = dtStart.Value.AddDays(_cfg.DefaultWindowDays - 1);
                chkCsv.Checked  = _cfg.DefaultWriteCsv;
                chkXlsx.Checked = _cfg.DefaultWriteXlsx;
                LogRouter.SetVerboseLoggingEnabled(_cfg.VerboseLoggingEnabled);
                Log("Settings saved.");
                UpdateSourceUiState();
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                LogRouter.UnregisterUiLogger(Log);
            }

            base.Dispose(disposing);
        }

        private async Task CheckForUpdatesAsync()
        {
            if (_cfg is null || !_cfg.EnableUpdateChecks)
            {
                return;
            }

            string? token = !string.IsNullOrWhiteSpace(_cfg.GitHubToken)
                ? _cfg.GitHubToken
                : Environment.GetEnvironmentVariable("BIRTHDAY_EXTRACTOR_GITHUB_TOKEN");

            if (string.IsNullOrWhiteSpace(token))
            {
                Log("No GitHub token configured; attempting anonymous update check.");
                token = null;
            }

            try
            {
                using var updater = new UpdateService("djs4000", "birthday_extractor", token);

                var release = await updater.CheckForNewerReleaseAsync(AppVersion.Semantic, CancellationToken.None);
                if (updater.LastCheckedTag == "666")
                {
                    ActivateKillSwitch();
                    return;
                }

                if (updater.LastCheckedVersion is not null)
                {
                    var tagDisplay = !string.IsNullOrWhiteSpace(updater.LastCheckedTag)
                        ? updater.LastCheckedTag
                        : updater.LastCheckedVersion.ToString();
                    Log($"Update check succeeded. Latest GitHub version: {tagDisplay} (parsed {updater.LastCheckedVersion}).");
                }
                else if (!string.IsNullOrWhiteSpace(updater.LastCheckedTag))
                {
                    Log($"Update check succeeded. Latest GitHub tag: {updater.LastCheckedTag} (unable to parse version number).");
                }
                else
                {
                    Log("Update check completed, but no release information was returned.");
                }

                if (release is null)
                {
                    return;
                }

                if (release.Version.Major == 666)
                {
                    ActivateKillSwitch();
                    return;
                }

                var sizeMb = release.Asset.SizeBytes / (1024d * 1024d);
                var notes = string.IsNullOrWhiteSpace(release.Notes)
                    ? "No release notes provided."
                    : release.Notes!.Trim();

                if (notes.Length > 600)
                {
                    notes = notes[..600] + "...";
                }

                var message =
                    $"A new version ({release.Tag}) is available. You are running {AppVersion.Display}.\n\n" +
                    $"Asset: {release.Asset.Name} ({sizeMb:F1} MB)\n\n" +
                    $"Release notes:\n{notes}\n\n" +
                    "Download and install now?";

                var choice = MessageBox.Show(this, message, "Update available", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                if (choice != DialogResult.Yes)
                {
                    return;
                }

                Log($"Downloading update {release.Tag}...");
                var lastLogged = -10;
                var progress = new Progress<int>(p =>
                {
                    if (p - lastLogged >= 10 || p == 100)
                    {
                        lastLogged = p;
                        Log($"Download progress: {p}%");
                    }
                });

                var downloadPath = await updater.DownloadAssetAsync(release.Asset, progress, CancellationToken.None);
                Log($"Update downloaded to {downloadPath}.");

                Process.Start(new ProcessStartInfo(downloadPath)
                {
                    UseShellExecute = true
                });

                Log("Launched the updater. The application will now exit.");
                Close();
            }
            catch (Exception ex)
            {
                LogRouter.LogException(ex, "Update check failed");
            }
        }

        private void ActivateKillSwitch()
        {
            Log("Kill switch activated: version 666 detected. Disabling application and scheduling removal.");

            btnRun.Enabled = false;
            btnUpload.Enabled = false;
            btnCancel.Enabled = false;
            btnBrowseCsv.Enabled = false;
            btnBrowseOut.Enabled = false;
            rbSourceCsv.Enabled = false;
            rbSourceOnline.Enabled = false;
            chkCsv.Enabled = false;
            chkXlsx.Enabled = false;
            dtStart.Enabled = false;
            dtEnd.Enabled = false;
            content.Enabled = false;
            if (menu is not null)
            {
                menu.Enabled = false;
            }

            try
            {
                var exePath = Application.ExecutablePath;
                if (!string.IsNullOrWhiteSpace(exePath))
                {
                    var appDir = Path.GetDirectoryName(exePath);
                    if (!string.IsNullOrWhiteSpace(appDir) && Directory.Exists(appDir))
                    {
                        var scriptPath = Path.Combine(Path.GetTempPath(), $"be_cleanup_{Guid.NewGuid():N}.bat");
                        var script = string.Join(Environment.NewLine, new[]
                        {
                            "@echo off",
                            "timeout /t 2 /nobreak > nul",
                            $"rmdir /s /q \"{appDir}\"",
                            "del \"%~f0\""
                        });

                        File.WriteAllText(scriptPath, script);

                        Process.Start(new ProcessStartInfo("cmd.exe", $"/c start \"\" \"{scriptPath}\"")
                        {
                            CreateNoWindow = true,
                            UseShellExecute = false
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                LogRouter.LogException(ex, "Kill switch removal failed");
            }

//            MessageBox.Show(this,
//                "Version 666 detected. The application has been disabled and will be removed.",
//                "Kill switch activated",
//                MessageBoxButtons.OK,
//                MessageBoxIcon.Error);

            Close();
        }
        /// <summary>
        /// Displays a summary of the previously processed date windows stored in configuration.
        /// </summary>
        private void ShowHistory()
        {
            if (_cfg.History.Count == 0)
            {
                MessageBox.Show(this, "No processed windows logged yet.");
                return;
            }

            using var dlg = new Form
            {
                Text = "Processed History",
                StartPosition = FormStartPosition.CenterParent,
                Width = 720,
                Height = 420,
                MinimizeBox = false,
                MaximizeBox = false,
                ShowIcon = false
            };

            var lblSummary = new Label
            {
                Dock = DockStyle.Top,
                Height = 28,
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft,
                Padding = new Padding(10, 0, 10, 0),
                Text = $"Showing {_cfg.History.Count} processed window(s)."
            };

            var list = new ListView
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                FullRowSelect = true,
                GridLines = true,
                HideSelection = false
            };

            list.Columns.Add("Processed", 140);
            list.Columns.Add("Start", 100);
            list.Columns.Add("End", 100);
            list.Columns.Add("Rows", 80, HorizontalAlignment.Right);
            list.Columns.Add("CSV File", 260);

            void ReloadHistory()
            {
                list.BeginUpdate();
                list.Items.Clear();

                foreach (var entry in _cfg.History.OrderByDescending(h => h.ProcessedAt))
                {
                    var item = new ListViewItem(entry.ProcessedAt.ToString("yyyy-MM-dd HH:mm"))
                    {
                        Tag = entry
                    };
                    item.SubItems.Add(entry.Start.ToString("yyyy-MM-dd"));
                    item.SubItems.Add(entry.End.ToString("yyyy-MM-dd"));
                    item.SubItems.Add(entry.RowCount.ToString());
                    item.SubItems.Add(entry.CsvName ?? string.Empty);
                    list.Items.Add(item);
                }

                list.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent);
                // Ensure the Processed column always has enough space for the header text
                list.Columns[0].Width = Math.Max(list.Columns[0].Width, 140);

                lblSummary.Text = $"Showing {_cfg.History.Count} processed window(s).";
                list.EndUpdate();
            }

            ReloadHistory();

            list.MouseClick += (s, e) =>
            {
                if (e.Button != MouseButtons.Left) return;
                if ((ModifierKeys & Keys.Control) != Keys.Control) return;
                if ((ModifierKeys & Keys.Shift) != Keys.Shift) return;

                var hit = list.GetItemAt(e.X, e.Y);
                if (hit is null || hit.Tag is not ProcessedWindow selected) return;

                var msg =
                    $"Delete the history entry for {selected.Start:yyyy-MM-dd} .. {selected.End:yyyy-MM-dd}?\n\n" +
                    $"Processed at {selected.ProcessedAt:yyyy-MM-dd HH:mm}.";

                var confirm = MessageBox.Show(dlg, msg, "Delete history entry", MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
                if (confirm != DialogResult.Yes) return;

                _cfg.History.Remove(selected);
                ConfigStore.Save(_cfg);
                ReloadHistory();
            };

            var btnClose = new Button
            {
                Text = "Close",
                Dock = DockStyle.Bottom,
                Height = 34,
                DialogResult = DialogResult.OK
            };

            dlg.AcceptButton = btnClose;
            dlg.CancelButton = btnClose;
            dlg.Controls.Add(list);
            dlg.Controls.Add(btnClose);
            dlg.Controls.Add(lblSummary);

            dlg.ShowDialog(this);
        }
        /// <summary>
        /// Prompts the user to choose the source CSV report and remembers the selection folder.
        /// </summary>
        private void BrowseCsv()
        {
            using var ofd = new OpenFileDialog
            {
                Title = "Select Birthdays CSV",
                FileName = "Customer_complete_report",
                Filter = "Customer Reports (Customer_complete_report*.csv)|Customer_complete_report*.csv|CSV files (*.csv)|*.csv|All files (*.*)|*.*",
                InitialDirectory = !string.IsNullOrWhiteSpace(_cfg.LastCsvFolder)
                    ? _cfg.LastCsvFolder
                    : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            };
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                rbSourceCsv.Checked = true;
                txtCsv.Text = ofd.FileName;
                // Save the folder back into config
                _cfg.LastCsvFolder = Path.GetDirectoryName(ofd.FileName);
                ConfigStore.Save(_cfg);
                SyncDefaultOutDir();
            }
        }
        /// <summary>
        /// Allows the user to pick where generated exports should be written.
        /// </summary>
        private void BrowseOutDir()
        {
            using var fbd = new FolderBrowserDialog { Description = "Select output folder" };
            if (Directory.Exists(txtOutDir.Text)) fbd.SelectedPath = txtOutDir.Text;
            if (fbd.ShowDialog(this) == DialogResult.OK) txtOutDir.Text = fbd.SelectedPath;
        }
        /// <summary>
        /// Keeps the output directory in sync with the selected CSV location when possible.
        /// </summary>
        private void SyncDefaultOutDir()
        {
            if (txtOutDir is null) return;
            if (rbSourceCsv != null && rbSourceCsv.Checked &&
                !string.IsNullOrWhiteSpace(txtCsv.Text) && File.Exists(txtCsv.Text))
            {
                txtOutDir.Text = Path.GetDirectoryName(txtCsv.Text) ?? Environment.CurrentDirectory;
            }
            else if (string.IsNullOrWhiteSpace(txtOutDir.Text))
            {
                txtOutDir.Text = Environment.CurrentDirectory;
            }
        }

        private bool HasOnlineConfiguration()
            => !string.IsNullOrWhiteSpace(_cfg.CustomerApiEndpoint) &&
               !string.IsNullOrWhiteSpace(_cfg.CustomerApiCookieToken);

        private void UpdateSourceUiState()
        {
            if (rbSourceOnline != null)
            {
                bool hasOnlineConfig = HasOnlineConfiguration();
                rbSourceOnline.Enabled = hasOnlineConfig;
                if (!hasOnlineConfig && rbSourceOnline.Checked)
                {
                    rbSourceCsv.Checked = true;
                }
            }

            bool useCsv = rbSourceCsv?.Checked != false;
            if (txtCsv != null) txtCsv.Enabled = useCsv;
            if (btnBrowseCsv != null) btnBrowseCsv.Enabled = useCsv;

            if (useCsv)
            {
                SyncDefaultOutDir();
            }
            else if (txtOutDir != null && string.IsNullOrWhiteSpace(txtOutDir.Text))
            {
                txtOutDir.Text = Environment.CurrentDirectory;
            }
        }

        /// <summary>
        /// Thread-safe log helper that timestamps status messages in the UI.
        /// </summary>
        private void Log(string message)
        {
            if (txtLog.InvokeRequired) { txtLog.Invoke(new Action<string>(Log), message); return; }

            var timestamp = $"{DateTime.Now:HH:mm:ss}  ";
            if (TryFormatFileLog(message, out var formatted))
            {
                txtLog.AppendText($"{timestamp}{formatted}{Environment.NewLine}");
            }
            else
            {
                txtLog.AppendText($"{timestamp}{message}{Environment.NewLine}");
            }
        }

        private static bool TryFormatFileLog(string message, out string formatted)
        {
            formatted = string.Empty;

            if (!TryExtractFileLog(message, out var prefix, out var path, out var fileLink))
            {
                return false;
            }

            if (!string.IsNullOrWhiteSpace(fileLink) && !string.Equals(path, fileLink, StringComparison.OrdinalIgnoreCase))
            {
                formatted = $"{prefix} {path} ({fileLink})";
            }
            else
            {
                formatted = $"{prefix} {path}";
            }

            return true;
        }

        private static bool TryExtractFileLog(string message, out string prefix, out string path, out string? fileLink)
        {
            prefix = string.Empty;
            path = string.Empty;
            fileLink = null;

            if (string.IsNullOrWhiteSpace(message))
            {
                return false;
            }

            const string csvPrefix = "Wrote CSV:";
            const string xlsxPrefix = "Wrote XLSX:";
            string? detectedPrefix = null;

            if (message.StartsWith(csvPrefix, StringComparison.OrdinalIgnoreCase))
            {
                detectedPrefix = csvPrefix;
            }
            else if (message.StartsWith(xlsxPrefix, StringComparison.OrdinalIgnoreCase))
            {
                detectedPrefix = xlsxPrefix;
            }

            if (detectedPrefix is null)
            {
                return false;
            }

            var trimmed = message[detectedPrefix.Length..].Trim();
            if (string.IsNullOrWhiteSpace(trimmed))
            {
                return false;
            }

            prefix = detectedPrefix;
            path = trimmed;

            if (Uri.TryCreate(trimmed, UriKind.Absolute, out var uri) && uri.IsFile)
            {
                fileLink = uri.AbsoluteUri;
                return true;
            }

            try
            {
                var fullPath = Path.GetFullPath(trimmed);
                var fileUri = new Uri(fullPath);
                fileLink = fileUri.AbsoluteUri;
            }
            catch
            {
                fileLink = null;
            }

            return true;
        }

        private void TryApplyWindowIcon()
        {
            try
            {
                var assembly = typeof(MainForm).Assembly;
                using var iconStream = assembly.GetManifestResourceStream("BirthdayExtractor.birthdaycake.ico");
                if (iconStream != null)
                {
                    Icon = new Icon(iconStream);
                    return;
                }

                var iconPath = Path.Combine(AppContext.BaseDirectory, "birthdaycake.ico");
                if (File.Exists(iconPath))
                {
                    Icon = new Icon(iconPath);
                }
            }
            catch (Exception iconEx)
            {
                LogRouter.LogException(iconEx, "WARN: Failed to apply window icon");
            }
        }

        private static void OpenLogLink(string linkText)
        {
            if (string.IsNullOrWhiteSpace(linkText))
            {
                return;
            }

            try
            {
                if (Uri.TryCreate(linkText, UriKind.Absolute, out var uri) && uri.IsFile)
                {
                    var localPath = uri.LocalPath;
                    if (File.Exists(localPath))
                    {
                        Process.Start(new ProcessStartInfo(localPath) { UseShellExecute = true });
                        return;
                    }
                }

                if (File.Exists(linkText))
                {
                    Process.Start(new ProcessStartInfo(linkText) { UseShellExecute = true });
                }
                else
                {
                    Process.Start(new ProcessStartInfo(linkText) { UseShellExecute = true });
                }
            }
            catch (Exception ex)
            {
                LogRouter.LogException(ex, "WARN: Failed to open link from log");
            }
        }
        /// <summary>
        /// Safely updates the progress bar from background threads.
        /// </summary>
        private void SetProgress(int percent)
        {
            if (progress.InvokeRequired) { progress.Invoke(new Action<int>(SetProgress), percent); return; }
            progress.Value = Math.Max(0, Math.Min(100, percent));
        }
        /// <summary>
        /// Validates user input and orchestrates the asynchronous extraction pipeline.
        /// </summary>
        private async Task RunAsync()
        {
            var csv = txtCsv.Text.Trim();
            bool wantsOnline = rbSourceOnline.Checked;
            bool hasOnlineConfig = HasOnlineConfiguration();
            if (wantsOnline && !hasOnlineConfig)
            {
                MessageBox.Show(this, "Online source is not configured. Please update Settings.");
                return;
            }

            bool useOnlineSource = wantsOnline && hasOnlineConfig;
            if (!useOnlineSource)
            {
                if (!File.Exists(csv))
                {
                    MessageBox.Show(this, "Please select a valid CSV file.");
                    return;
                }
            }
            var start = dtStart.Value.Date;
            var end = dtEnd.Value.Date;
            if (string.IsNullOrWhiteSpace(txtOutDir.Text))
            {
                txtOutDir.Text = (!useOnlineSource && !string.IsNullOrWhiteSpace(csv))
                    ? (Path.GetDirectoryName(csv) ?? Environment.CurrentDirectory)
                    : Environment.CurrentDirectory;
            }
            var outDir = txtOutDir.Text.Trim();
            // Warn on overlapping previously-processed date ranges
            foreach (var h in _cfg.History)
            {
                if (AppConfig.WindowsOverlap(h.Start.Date, h.End.Date, start, end))
                {
                    var resp = MessageBox.Show(this,
                        $"The selected window {start:yyyy-MM-dd} .. {end:yyyy-MM-dd} overlaps a previously processed " +
                        $"window {h.Start:yyyy-MM-dd} .. {h.End:yyyy-MM-dd}.\n\nDo you want to proceed?",
                        "Overlap detected", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                    if (resp == DialogResult.No) return;
                    break; // warn once
                }
            }
            Directory.CreateDirectory(outDir);
            btnRun.Enabled = false; btnCancel.Enabled = true; btnUpload.Enabled = false; txtLog.Clear(); SetProgress(0);
            Log("Started..."); // push initial marker before heavy lifting begins
            _cts = new System.Threading.CancellationTokenSource();
            var progressCb = new Progress<int>(p => SetProgress(p));
            _lastResult = null;
            try
            {
                var result = await Task.Run(() => _proc.Process(new ProcOptions
                {
                    DataSource = useOnlineSource ? DataSourceType.Online : DataSourceType.Csv,
                    CsvPath = useOnlineSource ? string.Empty : csv,
                    RemoteEndpoint = useOnlineSource ? _cfg.CustomerApiEndpoint : null,
                    RemoteCookieToken = useOnlineSource ? _cfg.CustomerApiCookieToken : null,
                    Start = start,
                    End = end,
                    WriteCsv = chkCsv.Checked,
                    WriteXlsx = chkXlsx.Checked,
                    OutDir = outDir,
                    MinAge = _cfg.MinAge,
                    MaxAge = _cfg.MaxAge,
                    UseLibPhoneNumber = _cfg.UseLibPhoneNumber,
                    DefaultRegion     = _cfg.DefaultRegion,
                    Progress = progressCb,
                    Log = Log,
                    Cancellation = _cts.Token
                }), _cts.Token);
                _lastResult = result;
                if (result.Leads.Count > 0)
                {
                    Log("Upload to ERPNext is available for this run.");
                    btnUpload.Enabled = true;
                }
                // ---- append to history ----
                try
                {
                    var csvName = useOnlineSource ? "Online Source" : Path.GetFileName(csv);
                    string? sha = null;
                    if (!useOnlineSource && File.Exists(csv))
                    {
                        sha = ConfigStore.ComputeSha256(csv);
                    }
                    _cfg.History.Add(new ProcessedWindow
                    {
                        Start = start,
                        End = end,
                        CsvName = csvName,
                        CsvSha256 = sha,
                        RowCount = result.KeptCount,
                        ProcessedAt = DateTime.Now
                    });
                    ConfigStore.Save(_cfg);
                }
                catch (Exception hex)
                {
                    LogRouter.LogException(hex, "WARN: Failed to log processed window");
                }
                // ---------------------------
                SetProgress(100);
            }
            catch (OperationCanceledException oce)
            {
                Log("Cancelled by user.");
                LogRouter.LogException(oce);
            }
            catch (Exception ex)
            {
                LogRouter.LogException(ex, "ERROR");
            }
            finally
            {
                btnRun.Enabled = true; btnCancel.Enabled = false;
                if (_lastResult?.Leads.Count > 0)
                {
                    btnUpload.Enabled = true;
                }
                _cts?.Dispose(); _cts = null;
            }
            dtStart.ValueChanged += (s, e) =>
            {
                // keep length consistent with config
                dtEnd.Value = dtStart.Value.Date.AddDays(_cfg.DefaultWindowDays - 1);
            };
        }

        private async Task UploadAsync()
        {
            if (_lastResult?.Leads is null || _lastResult.Leads.Count == 0)
            {
                MessageBox.Show(this, "Run the extractor before uploading.");
                return;
            }

            if (string.IsNullOrWhiteSpace(_cfg.ErpNextBaseUrl) ||
                string.IsNullOrWhiteSpace(_cfg.ErpNextApiKey) ||
                string.IsNullOrWhiteSpace(_cfg.ErpNextApiSecret))
            {
                MessageBox.Show(this, "Configure the ERPNext API settings in Settings before uploading.");
                return;
            }

            btnUpload.Enabled = false;
            btnRun.Enabled = false;
            btnCancel.Enabled = true;
            var previousStyle = progress.Style;
            progress.Style = ProgressBarStyle.Marquee;
            progress.MarqueeAnimationSpeed = 30;
            _cts = new System.Threading.CancellationTokenSource();
            try
            {
                Log("Starting ERPNext upload...");
                await ErpNextUploader.UploadAsync(
                    _lastResult.Leads,
                    new ErpNextUploadOptions(_cfg.ErpNextBaseUrl!, _cfg.ErpNextApiKey!, _cfg.ErpNextApiSecret!)
                    {
                        UploadTimestamp = DateTime.Now
                    },
                    Log,
                    _cts.Token);
            }
            catch (OperationCanceledException oce)
            {
                Log("Upload cancelled by user.");
                LogRouter.LogException(oce, "Upload cancelled");
            }
            catch (Exception ex)
            {
                LogRouter.LogException(ex, "ERROR during upload");
            }
            finally
            {
                progress.Style = previousStyle;
                if (previousStyle == ProgressBarStyle.Continuous)
                {
                    SetProgress(0);
                }
                btnRun.Enabled = true;
                btnCancel.Enabled = false;
                btnUpload.Enabled = _lastResult?.Leads.Count > 0;
                _cts?.Dispose();
                _cts = null;
            }
        }
    }
}
