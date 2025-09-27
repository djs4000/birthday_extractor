using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
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
        private TextBox txtLog = null!;
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
            catch
            {
                _cfg = new AppConfig();
            }
            // sanity defaults if someone hand-edited config
            if (_cfg.DefaultWindowDays <= 0) _cfg.DefaultWindowDays = 7;
            // 2) Form shell: establish window chrome before wiring controls
            Text = $"Birthday Extractor v{AppVersion.Display}";
            Width = 820; Height = 600;
            StartPosition = FormStartPosition.CenterScreen;

            // 3) Build UI
            InitializeMenu();
            InitializeContentPanel();

            // 4) Defaults pulled from config to pre-populate the form
            dtStart.Value = DateTime.Today.AddDays(_cfg.DefaultStartOffsetDays);
            dtEnd.Value   = dtStart.Value.AddDays(_cfg.DefaultWindowDays - 1);
            chkCsv.Checked  = _cfg.DefaultWriteCsv;
            chkXlsx.Checked = _cfg.DefaultWriteXlsx;
            dtStart.ValueChanged += (s, e) => dtEnd.Value = dtStart.Value.Date.AddDays(_cfg.DefaultWindowDays - 1);
            txtCsv.TextChanged += (s, e) => SyncDefaultOutDir();
            SyncDefaultOutDir();
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
            content = new Panel { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(10) };
            Controls.Add(content);
            // 5) Build your controls (labels, inputs, and action buttons)
            var lblCsv = new Label { Left = 20, Top = 40, Width = 120, Text = "CSV File:" };
            txtCsv = new TextBox { Left = 140, Top = 36, Width = 540 };
            btnBrowseCsv = new Button { Left = 690, Top = 34, Width = 90, Text = "Browse..." };
            btnBrowseCsv.Click += (s, e) => BrowseCsv();
            var lblStart = new Label { Left = 20, Top = 70, Width = 120, Text = "Start Date:" };
            dtStart = new DateTimePicker { Left = 140, Top = 66, Width = 200, Format = DateTimePickerFormat.Custom, CustomFormat = "yyyy-MM-dd" };
            dtEnd = new DateTimePicker { Left = 440, Top = 66, Width = 200, Format = DateTimePickerFormat.Custom, CustomFormat = "yyyy-MM-dd" };
            var lblEnd = new Label { Left = 360, Top = 70, Width = 60, Text = "End Date:" };
            chkCsv = new CheckBox { Left = 140, Top = 96, Width = 80, Text = "CSV" };
            chkXlsx = new CheckBox { Left = 230, Top = 96, Width = 80, Text = "XLSX" };
            var lblOut = new Label { Left = 20, Top = 136, Width = 120, Text = "Output Folder:" };
            txtOutDir = new TextBox { Left = 140, Top = 132, Width = 540 };
            btnBrowseOut = new Button { Left = 690, Top = 130, Width = 90, Text = "Browse..." };
            btnBrowseOut.Click += (s, e) => BrowseOutDir();
            btnRun = new Button { Left = 140, Top = 172, Width = 120, Text = "Run" };
            btnCancel = new Button { Left = 270, Top = 172, Width = 120, Text = "Cancel", Enabled = false };
            btnUpload = new Button { Left = 400, Top = 172, Width = 150, Text = "Upload to ERPNext", Enabled = false };
            btnRun.Click += async (s, e) => await RunAsync();
            btnCancel.Click += (s, e) => _cts?.Cancel();
            btnUpload.Click += async (s, e) => await UploadAsync();
            progress = new ProgressBar { Left = 20, Top = 212, Width = 760, Height = 18, Style = ProgressBarStyle.Continuous, Minimum = 0, Maximum = 100, Value = 0 };
            txtLog = new TextBox { Left = 20, Top = 242, Width = 760, Height = 300, Multiline = true, ScrollBars = ScrollBars.Vertical, ReadOnly = true };
            // 6) Add to content panel (not the form)
            content.Controls.AddRange(new Control[] {
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
                Log("Settings saved.");
            }
        }

        private async Task CheckForUpdatesAsync()
        {
            if (_cfg is null || !_cfg.EnableUpdateChecks)
            {
                return;
            }

            var token = !string.IsNullOrWhiteSpace(_cfg.GitHubToken)
                ? _cfg.GitHubToken
                : Environment.GetEnvironmentVariable("BIRTHDAY_EXTRACTOR_GITHUB_TOKEN");

            if (string.IsNullOrWhiteSpace(token))
            {
                Log("Update check skipped: no GitHub token configured.");
                return;
            }

            try
            {
                using var updater = new UpdateService("djs4000", "birthday_extractor", token);
                var release = await updater.CheckForNewerReleaseAsync(AppVersion.Semantic, CancellationToken.None);
                if (release is null)
                {
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
                Log("Update check failed: " + ex.Message);
            }
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
            txtOutDir.Text = (!string.IsNullOrWhiteSpace(txtCsv.Text) && File.Exists(txtCsv.Text))
                ? (Path.GetDirectoryName(txtCsv.Text) ?? Environment.CurrentDirectory)
                : Environment.CurrentDirectory;
        }
        /// <summary>
        /// Thread-safe log helper that timestamps status messages in the UI.
        /// </summary>
        private void Log(string message)
        {
            if (txtLog.InvokeRequired) { txtLog.Invoke(new Action<string>(Log), message); return; }
            txtLog.AppendText($"{DateTime.Now:HH:mm:ss}  {message}{Environment.NewLine}");
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
            if (!File.Exists(csv)) { MessageBox.Show(this, "Please select a valid CSV file."); return; }
            var start = dtStart.Value.Date;
            var end = dtEnd.Value.Date;
            if (string.IsNullOrWhiteSpace(txtOutDir.Text)) txtOutDir.Text = Path.GetDirectoryName(csv) ?? Environment.CurrentDirectory;
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
                    CsvPath = csv,
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
                Log($"Done. Kept {result.KeptCount} rows.");
                if (result.CsvPath is not null) Log($"CSV : {result.CsvPath}");
                if (result.XlsxPath is not null) Log($"XLSX: {result.XlsxPath}");
                _lastResult = result;
                if (result.Leads.Count > 0)
                {
                    Log("Upload to ERPNext is available for this run.");
                    btnUpload.Enabled = true;
                }
                // ---- append to history ----
                try
                {
                    var csvName = Path.GetFileName(csv);
                    var sha = ConfigStore.ComputeSha256(csv);
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
                    Log("WARN: Failed to log processed window: " + hex.Message);
                }
                // ---------------------------
                SetProgress(100);
            }
            catch (OperationCanceledException)
            {
                Log("Cancelled by user.");
            }
            catch (Exception ex)
            {
                Log("ERROR: " + ex.Message);
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
            catch (OperationCanceledException)
            {
                Log("Upload cancelled by user.");
            }
            catch (Exception ex)
            {
                Log("ERROR during upload: " + ex.Message);
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
