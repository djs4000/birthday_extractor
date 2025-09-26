using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BirthdayExtractor
{
    public class MainForm : Form
    {
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
        private ProgressBar progress = null!;
        private TextBox txtLog = null!;

        private readonly Processing _proc = new();
        private System.Threading.CancellationTokenSource? _cts;
        private AppConfig _cfg = null!;
        private MenuStrip menu = null!;
        private ToolStripMenuItem miSettings = null!;
        private ToolStripMenuItem miHistory  = null!;


        public MainForm()
        {
            // 1) Load config FIRST, with a safe fallback
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

            // 2) Form shell
            Text = "Birthday Extractor v0.5";
            Width = 820; Height = 600;
            StartPosition = FormStartPosition.CenterScreen;

            // 3) Menu (Dock Top)
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

            // 4) Content panel (Dock Fill) – all inputs go here
            content = new Panel { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(10) };
            Controls.Add(content);

            // 5) Build your controls (same as before)
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
            btnRun.Click += async (s, e) => await RunAsync();
            btnCancel.Click += (s, e) => _cts?.Cancel();

            progress = new ProgressBar { Left = 20, Top = 212, Width = 760, Height = 18, Style = ProgressBarStyle.Continuous, Minimum = 0, Maximum = 100, Value = 0 };
            txtLog = new TextBox { Left = 20, Top = 242, Width = 760, Height = 300, Multiline = true, ScrollBars = ScrollBars.Vertical, ReadOnly = true };

            // 6) Add to content panel (not the form)
            content.Controls.AddRange(new Control[] {
                lblCsv, txtCsv, btnBrowseCsv,
                lblStart, dtStart,
                lblEnd, dtEnd,
                chkCsv, chkXlsx,
                lblOut, txtOutDir, btnBrowseOut,
                btnRun, btnCancel,
                progress, txtLog
            });

            // 7) Defaults
            dtStart.Value = DateTime.Today.AddDays(_cfg.DefaultStartOffsetDays);
            dtEnd.Value   = dtStart.Value.AddDays(_cfg.DefaultWindowDays - 1);
            chkCsv.Checked  = _cfg.DefaultWriteCsv;
            chkXlsx.Checked = _cfg.DefaultWriteXlsx;

            dtStart.ValueChanged += (s, e) => dtEnd.Value = dtStart.Value.Date.AddDays(_cfg.DefaultWindowDays - 1);

            txtCsv.TextChanged += (s, e) => SyncDefaultOutDir();
            SyncDefaultOutDir();
        }
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

        private void ShowHistory()
        {
            if (_cfg.History.Count == 0) { MessageBox.Show(this, "No processed windows logged yet."); return; }
            var lines = _cfg.History
                .OrderByDescending(h => h.ProcessedAt)
                .Select(h => $"{h.ProcessedAt:yyyy-MM-dd HH:mm}  {h.Start:yyyy-MM-dd} → {h.End:yyyy-MM-dd}  rows={h.RowCount}  csv={h.CsvName}")
                .ToArray();
            MessageBox.Show(this, string.Join(Environment.NewLine, lines), "Processed History",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

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

        private void BrowseOutDir()
        {
            using var fbd = new FolderBrowserDialog { Description = "Select output folder" };
            if (Directory.Exists(txtOutDir.Text)) fbd.SelectedPath = txtOutDir.Text;
            if (fbd.ShowDialog(this) == DialogResult.OK) txtOutDir.Text = fbd.SelectedPath;
        }

        private void SyncDefaultOutDir()
        {
            txtOutDir.Text = (!string.IsNullOrWhiteSpace(txtCsv.Text) && File.Exists(txtCsv.Text))
                ? (Path.GetDirectoryName(txtCsv.Text) ?? Environment.CurrentDirectory)
                : Environment.CurrentDirectory;
        }

        private void Log(string message)
        {
            if (txtLog.InvokeRequired) { txtLog.Invoke(new Action<string>(Log), message); return; }
            txtLog.AppendText($"{DateTime.Now:HH:mm:ss}  {message}{Environment.NewLine}");
        }

        private void SetProgress(int percent)
        {
            if (progress.InvokeRequired) { progress.Invoke(new Action<int>(SetProgress), percent); return; }
            progress.Value = Math.Max(0, Math.Min(100, percent));
        }

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
            btnRun.Enabled = false; btnCancel.Enabled = true; txtLog.Clear(); SetProgress(0);
            Log("Started...");

            _cts = new System.Threading.CancellationTokenSource();
            var progressCb = new Progress<int>(p => SetProgress(p));

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
                _cts?.Dispose(); _cts = null;
            }
            dtStart.ValueChanged += (s, e) =>
            {
                // keep length consistent with config
                dtEnd.Value = dtStart.Value.Date.AddDays(_cfg.DefaultWindowDays - 1);
            };
        }
    }
}
