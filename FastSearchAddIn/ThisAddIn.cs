using System;
using System.Windows.Forms;
using FastSearchAddIn.Core;
using FastSearchAddIn.UI;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools;
using NLog;

namespace FastSearchAddIn
{
    /// <summary>
    /// Entry point for the Outlook VSTO Add-in.
    /// Initialises the Lucene index and registers the WPF task pane.
    /// </summary>
    public partial class ThisAddIn
    {
        private static readonly Logger Log = LogManager.GetCurrentClassLogger();

        // The custom task pane host (VSTO wrapper around the WPF UserControl).
        private CustomTaskPane _searchTaskPane;

        // Core services – shared across the add-in lifetime.
        internal IndexManager IndexManager { get; private set; }
        internal SearchProvider SearchProvider { get; private set; }
        internal ExportManager ExportManager { get; private set; }

        /// <summary>
        /// Fires when Outlook loads the add-in.
        /// </summary>
        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            try
            {
                AppSettings.EnsureDirectories();
                ConfigureLogging();

                Log.Info("FastSearchAddIn starting up.");

                // Instantiate core services.
                IndexManager = new IndexManager();
                SearchProvider = new SearchProvider(IndexManager);
                ExportManager = new ExportManager();

                // Create the WPF task pane.
                var wpfControl = new SearchTaskPane(IndexManager, SearchProvider, ExportManager);
                var host = new System.Windows.Forms.Integration.ElementHost
                {
                    Child = wpfControl,
                    Dock = DockStyle.Fill
                };

                // Wrap inside a WinForms UserControl so VSTO can host it.
                var winFormsWrapper = new UserControl();
                winFormsWrapper.Controls.Add(host);

                _searchTaskPane = CustomTaskPanes.Add(winFormsWrapper, "Fast Search");
                _searchTaskPane.Width = 380;
                _searchTaskPane.Visible = true;

                // Kick off background indexing.
                IndexManager.StartIndexingAsync(Application);

                Log.Info("FastSearchAddIn started successfully.");
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Error during add-in startup.");
                MessageBox.Show(
                    $"FastSearchAddIn failed to start: {ex.Message}\n\nSee log for details.",
                    "FastSearchAddIn – Startup Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Fires when Outlook shuts down the add-in.
        /// </summary>
        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            try
            {
                Log.Info("FastSearchAddIn shutting down.");
                IndexManager?.Dispose();
                SearchProvider?.Dispose();
                // ExportManager has no disposable resources – nothing to release here.
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Error during add-in shutdown.");
            }
        }

        // -----------------------------------------------------------------------
        // VSTO generated code – do not modify.
        // -----------------------------------------------------------------------
        #region VSTO generated code
        private void InternalStartup()
        {
            Startup += new EventHandler(ThisAddIn_Startup);
            Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }
        #endregion

        // -----------------------------------------------------------------------
        // Helpers
        // -----------------------------------------------------------------------

        /// <summary>
        /// Configure NLog to write to %LocalAppData%\FastSearchAddIn\Logs\addin.log.
        /// </summary>
        private static void ConfigureLogging()
        {
            var config = new NLog.Config.LoggingConfiguration();
            var fileTarget = new NLog.Targets.FileTarget("logfile")
            {
                FileName = System.IO.Path.Combine(AppSettings.LogsFolder, "addin.log"),
                Layout = "${longdate} [${level:uppercase=true}] ${logger} – ${message} ${exception:format=tostring}",
                ArchiveAboveSize = 5_000_000,   // 5 MB
                MaxArchiveFiles = 3,
                ArchiveNumbering = NLog.Targets.ArchiveNumberingMode.Rolling
            };
            config.AddRule(NLog.LogLevel.Debug, NLog.LogLevel.Fatal, fileTarget);
            LogManager.Configuration = config;
        }
    }
}
