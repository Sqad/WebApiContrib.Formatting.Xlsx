using System.Data;

namespace SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Views.Formatted
{
    internal class FormattedExcelDataRow : ExcelDataRow
    {
        public FormattedExcelDataRow(DataRow dataRow)
            : base(dataRow)
        {
            IsHeader = this["header"] != null && bool.Parse((string) this["header"]);
        }

        public bool IsHeader { get; }
    }
}