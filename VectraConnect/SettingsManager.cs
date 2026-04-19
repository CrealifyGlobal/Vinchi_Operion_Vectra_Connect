using Microsoft.Win32;

namespace VectraConnect
{
    /// <summary>
    /// Persists add-in settings to HKCU registry so they survive across sessions.
    /// </summary>
    public static class SettingsManager
    {
        private const string RegKey = @"Software\VectraConnect";

        public static string OutputFolder
        {
            get => Read("OutputFolder", "");
            set => Write("OutputFolder", value);
        }

        public static bool IncludeCsv
        {
            get => Read("IncludeCsv", "1") == "1";
            set => Write("IncludeCsv", value ? "1" : "0");
        }

        // ── Helpers ───────────────────────────────────────────────────────

        private static string Read(string name, string defaultValue)
        {
            using (var key = Registry.CurrentUser.OpenSubKey(RegKey))
                return key?.GetValue(name, defaultValue)?.ToString() ?? defaultValue;
        }

        private static void Write(string name, string value)
        {
            using (var key = Registry.CurrentUser.CreateSubKey(RegKey))
                key?.SetValue(name, value);
        }
    }
}
