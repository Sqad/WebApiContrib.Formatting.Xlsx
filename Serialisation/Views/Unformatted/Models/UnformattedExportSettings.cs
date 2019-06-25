using System.Data;
using System.Linq;

namespace WebApiContrib.Formatting.Xlsx.Serialisation.Views.Unformatted.Models
{
    internal class UnformattedExportSettings
    {
        public UnformattedExportSettings(DataTable settingsTable)
        {
            UseNewVersion = (bool?) GetSetting(settingsTable, "UseNewVersion") ?? false;
            ExcelLink = (string) GetSetting(settingsTable, "ExcelLink");
            UseEmbeddedLogin = (bool?) GetSetting(settingsTable, "ExcelLink") ?? false;
            TokenPageLink = (string) GetSetting(settingsTable, "TokenPageLink");
            LoginPageLink = (string) GetSetting(settingsTable, "LoginPageLink");
        }

        public bool UseNewVersion { get; }
        public string ExcelLink { get; }
        public bool UseEmbeddedLogin { get; }
        public string TokenPageLink { get; }
        public string LoginPageLink { get; }

        private static object GetSetting(DataTable settingsTable, string key)
        {
            return settingsTable.Select($"key = '{key}'").FirstOrDefault()?["value"];
        }
    }
}