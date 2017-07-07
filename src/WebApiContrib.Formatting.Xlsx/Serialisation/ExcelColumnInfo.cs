using WebApiContrib.Formatting.Xlsx.Attributes;

namespace WebApiContrib.Formatting.Xlsx.Serialisation
{
    /// <summary>
    /// Formatting information for an Excel column based on attribute values specified on a class.
    /// </summary>
    public class ExcelColumnInfo
    {
        public string PropertyName { get; set; }
        public ExcelColumnAttribute ExcelColumnAttribute { get; set; }
        public string FormatString { get; set; }
        public string Header { get; set; }

        

        public string ExcelNumberFormat
        {
            get { return ExcelColumnAttribute != null ? ExcelColumnAttribute.NumberFormat : null; }
        }

        public bool IsExcelHeaderDefined
        {
            get { return ExcelColumnAttribute != null && ExcelColumnAttribute.Header != null; }
        }

        public ExcelColumnInfo(string propertyName, ExcelColumnAttribute excelAttribute = null, string formatString = null)
        {
            PropertyName = propertyName;
            ExcelColumnAttribute = excelAttribute;
            FormatString = formatString;
            Header = IsExcelHeaderDefined ? ExcelColumnAttribute.Header : propertyName;
        }
    }
}
