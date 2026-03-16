using System;
using System.IO;
using Newtonsoft.Json;

namespace FastSearchAddIn.Core
{
    /// <summary>
    /// Centralised paths and persistent settings stored in
    /// %LocalAppData%\FastSearchAddIn\settings.json.
    /// </summary>
    public static class AppSettings
    {
        // -----------------------------------------------------------------------
        // Well-known directories
        // -----------------------------------------------------------------------
        public static readonly string BaseFolder =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "FastSearchAddIn");

        public static readonly string IndexFolder   = Path.Combine(BaseFolder, "Index");
        public static readonly string LogsFolder    = Path.Combine(BaseFolder, "Logs");
        public static readonly string SettingsFile  = Path.Combine(BaseFolder, "settings.json");

        // -----------------------------------------------------------------------
        // Persistent settings
        // -----------------------------------------------------------------------

        /// <summary>Settings loaded from / saved to disk.</summary>
        private static Settings _current;

        public static Settings Current
        {
            get
            {
                if (_current == null) Load();
                return _current;
            }
        }

        public static void EnsureDirectories()
        {
            Directory.CreateDirectory(BaseFolder);
            Directory.CreateDirectory(IndexFolder);
            Directory.CreateDirectory(LogsFolder);
        }

        public static void Load()
        {
            if (File.Exists(SettingsFile))
            {
                try
                {
                    string json = File.ReadAllText(SettingsFile);
                    _current = JsonConvert.DeserializeObject<Settings>(json) ?? new Settings();
                    return;
                }
                catch { /* fall through to defaults */ }
            }
            _current = new Settings();
        }

        public static void Save()
        {
            try
            {
                string json = JsonConvert.SerializeObject(_current, Formatting.Indented);
                File.WriteAllText(SettingsFile, json);
            }
            catch { /* best-effort */ }
        }
    }

    /// <summary>User-configurable preferences.</summary>
    public class Settings
    {
        /// <summary>
        /// Comma-separated list of Outlook folder entry IDs to index.
        /// Empty means "index all folders".
        /// </summary>
        public string IndexedFolderEntryIds { get; set; } = string.Empty;

        /// <summary>Debounce delay (ms) between keystrokes and search execution.</summary>
        public int SearchDebounceMs { get; set; } = 300;

        /// <summary>Maximum number of search results to display.</summary>
        public int MaxResults { get; set; } = 200;

        /// <summary>Whether to also index email body text (slower to index).</summary>
        public bool IndexBody { get; set; } = true;
    }
}
