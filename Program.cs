using System;
using System.Windows.Forms;

namespace BirthdayExtractor
{
    /// <summary>
    /// Entry point for the Birthday Extractor Windows Forms application.
    /// Responsible for bootstrapping WinForms defaults and the main UI.
    /// </summary>
    internal static class Program
    {
        /// <summary>
        /// Boots the UI thread, applies visual/high DPI settings, and opens <see cref="MainForm"/>.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetHighDpiMode(HighDpiMode.SystemAware);
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }
}