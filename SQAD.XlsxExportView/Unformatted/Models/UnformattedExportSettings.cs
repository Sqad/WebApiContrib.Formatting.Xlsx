using System.Data;
using System.Linq;

namespace SQAD.XlsxExportView.Unformatted.Models
{
    internal class UnformattedExportSettings
    {
        public UnformattedExportSettings(DataTable settingsTable)
        {
            UseNewVersion = GetBoolSetting(settingsTable, "UseNewVersion");
            ExcelLink = GetStringSetting(settingsTable, "ExcelLink");
            UseEmbeddedLogin = GetBoolSetting(settingsTable, "UseEmbeddedLogin");
            TokenPageLink = GetStringSetting(settingsTable, "TokenPageLink");
            LoginPageLink = GetStringSetting(settingsTable, "LoginPageLink");
        }

        public bool UseNewVersion { get; }
        public string ExcelLink { get; }
        public bool UseEmbeddedLogin { get; }
        public string TokenPageLink { get; }
        public string LoginPageLink { get; }

        private static string GetStringSetting(DataTable settingsTable, string key)
        {
            return (string) settingsTable.Select($"key = '{key}'").FirstOrDefault()?["value"];
        }

        private static bool GetBoolSetting(DataTable settingsTable, string key)
        {
            var stringSettings = GetStringSetting(settingsTable, key);
            return bool.TryParse(stringSettings, out var result) && result;
        }
    }
}