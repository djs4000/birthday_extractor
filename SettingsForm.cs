// SettingsForm.cs
using System;
using System.Windows.Forms;
namespace BirthdayExtractor
{
    /// <summary>
    /// Secondary dialog for adjusting defaults and advanced options.
    /// </summary>
    public sealed class SettingsForm : Form
    {
        private NumericUpDown numOffset = null!;
        private NumericUpDown numWindow = null!;
        private NumericUpDown numMinAge = null!;
        private NumericUpDown numMaxAge = null!;
        private CheckBox chkCsv = null!;
        private CheckBox chkXlsx = null!;
        private CheckBox chkUseLibPhone = null!;   // <— add this
        private TextBox txtWebhookUrl = null!;
        private TextBox txtWebhookAuth = null!;
        private Button btnSave = null!;
        private Button btnCancel = null!;
        private readonly AppConfig _cfg;
        /// <summary>
        /// Builds the settings dialog around an existing configuration object.
        /// </summary>
        public SettingsForm(AppConfig cfg)
        {
            _cfg = cfg;
            Text = "Settings";
            Width = 520; Height = 360;
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false; MinimizeBox = false;
            var y = 20; // Layout cursor to stack controls vertically
            Controls.Add(new Label { Left = 20, Top = y, Width = 220, Text = "Default start offset (days from today):" });
            numOffset = new NumericUpDown { Left = 260, Top = y - 2, Width = 200, Minimum = 0, Maximum = 365, Value = _cfg.DefaultStartOffsetDays }; y += 30;
            Controls.Add(new Label { Left = 20, Top = y, Width = 220, Text = "Default window length (days):" });
            numWindow = new NumericUpDown { Left = 260, Top = y - 2, Width = 200, Minimum = 1, Maximum = 60, Value = _cfg.DefaultWindowDays }; y += 30;
            Controls.Add(new Label { Left = 20, Top = y, Width = 220, Text = "Age range (min / max):" });
            numMinAge = new NumericUpDown { Left = 260, Top = y - 2, Width = 90, Minimum = 0, Maximum = 120, Value = _cfg.MinAge };
            numMaxAge = new NumericUpDown { Left = 370, Top = y - 2, Width = 90, Minimum = 0, Maximum = 120, Value = _cfg.MaxAge }; y += 30;
            chkCsv  = new CheckBox { Left = 260, Top = y, Width = 80, Text = "CSV", Checked = _cfg.DefaultWriteCsv };
            chkXlsx = new CheckBox { Left = 340, Top = y, Width = 80, Text = "XLSX", Checked = _cfg.DefaultWriteXlsx }; y += 30;
            Controls.Add(chkCsv);
            Controls.Add(chkXlsx);
            // Advanced phone parsing (libphonenumber)
            chkUseLibPhone = new CheckBox {
                Left = 260, Top = y, Width = 240,
                Text = "Use libphonenumber (advanced)",
                Checked = _cfg.UseLibPhoneNumber
            };
            Controls.Add(new Label { Left = 20, Top = y + 2, Width = 220, Text = "Phone parsing:" });
            Controls.Add(chkUseLibPhone);
            y += 35;
            Controls.Add(new Label { Left = 20, Top = y, Width = 220, Text = "Webhook URL (future):" });
            txtWebhookUrl = new TextBox { Left = 260, Top = y - 4, Width = 200, Text = _cfg.WebhookUrl ?? "" }; y += 30;
            Controls.Add(new Label { Left = 20, Top = y, Width = 220, Text = "Webhook Auth header (future):" });
            txtWebhookAuth = new TextBox { Left = 260, Top = y - 4, Width = 200, Text = _cfg.WebhookAuthHeader ?? "" }; y += 40;
            btnSave = new Button { Left = 260, Top = y, Width = 90, Text = "Save" };
            btnCancel = new Button { Left = 370, Top = y, Width = 90, Text = "Cancel" };
            btnSave.Click += (s, e) => SaveAndClose();
            btnCancel.Click += (s, e) => DialogResult = DialogResult.Cancel;
            Controls.Add(btnSave);
            Controls.Add(btnCancel);
        }
        /// <summary>
        /// Applies form values back onto the config object and persists them.
        /// </summary>
        private void SaveAndClose()
        {
            _cfg.DefaultStartOffsetDays = (int)numOffset.Value;
            _cfg.DefaultWindowDays      = (int)numWindow.Value;
            _cfg.MinAge                 = (int)numMinAge.Value;
            _cfg.MaxAge                 = (int)numMaxAge.Value;
            _cfg.DefaultWriteCsv        = chkCsv.Checked;
            _cfg.DefaultWriteXlsx       = chkXlsx.Checked;
            _cfg.WebhookUrl             = string.IsNullOrWhiteSpace(txtWebhookUrl.Text) ? null : txtWebhookUrl.Text.Trim();
            _cfg.WebhookAuthHeader      = string.IsNullOrWhiteSpace(txtWebhookAuth.Text) ? null : txtWebhookAuth.Text.Trim();
            _cfg.UseLibPhoneNumber      = chkUseLibPhone.Checked;
            ConfigStore.Save(_cfg);
            DialogResult = DialogResult.OK;
        }
    }
}
