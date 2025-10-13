using System;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BirthdayExtractor
{
    /// <summary>
    /// Temporary debug helpers for manual update and kill switch testing.
    /// Segregated in its own partial so it can be removed before merging to main.
    /// </summary>
    public partial class MainForm
    {
        private ToolStripMenuItem? _miDebug;
        private ToolStripMenuItem? _miDebugTriggerUpdate;
        private ToolStripMenuItem? _miDebugTriggerKillSwitch;

        partial void InitializeDebugMenuItems(MenuStrip menu)
        {
            _miDebug = new ToolStripMenuItem("Debug")
            {
                ToolTipText = "Temporary helpers for exercising updater and kill switch logic"
            };

            _miDebugTriggerUpdate = new ToolStripMenuItem("Check for Updates Now", null, HandleDebugUpdateClick)
            {
                ShortcutKeys = Keys.Control | Keys.U
            };

            _miDebugTriggerKillSwitch = new ToolStripMenuItem("Activate Kill Switch", null, HandleDebugKillSwitchClick)
            {
                ShortcutKeys = Keys.Control | Keys.K
            };

            _miDebug.DropDownItems.Add(_miDebugTriggerUpdate);
            _miDebug.DropDownItems.Add(new ToolStripSeparator());
            _miDebug.DropDownItems.Add(_miDebugTriggerKillSwitch);

            menu.Items.Add(_miDebug);
        }

        private async void HandleDebugUpdateClick(object? sender, EventArgs e)
        {
            await TriggerManualUpdateAsync();
        }

        private async Task TriggerManualUpdateAsync()
        {
            try
            {
                Log("[DEBUG] Manual update check triggered from debug menu.");
                await CheckForUpdatesAsync(ignoreConfigSettings: true);
            }
            catch (Exception ex)
            {
                LogRouter.LogException(ex, "DEBUG: Manual update check failed");
            }
        }

        private void HandleDebugKillSwitchClick(object? sender, EventArgs e)
        {
            var confirm = MessageBox.Show(this,
                "Trigger the kill switch logic? This will disable the UI and schedule application removal.",
                "Confirm Kill Switch",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);

            if (confirm == DialogResult.Yes)
            {
                Log("[DEBUG] Manual kill switch triggered from debug menu.");
                ActivateKillSwitch();
            }
        }
    }
}
