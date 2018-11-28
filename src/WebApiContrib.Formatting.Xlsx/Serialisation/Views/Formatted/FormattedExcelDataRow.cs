using System;
using System.Data;
using System.Globalization;

namespace SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Views.Formatted
{
    internal class FormattedExcelDataRow : ExcelDataRow
    {
        public FormattedExcelDataRow(DataRow dataRow)
            : base(dataRow)
        {
            IsHeader = this["header"] != null && bool.Parse((string)this["header"]);
        }

        public bool IsHeader { get; }
        
        //note: DataTable values from ViewAPI already formatted, but Excel don't recognize it
        protected override object ParseValue(string columnName)
        {
            var value = (string) this[columnName];

            if (value == null)
            {
                return null;
            }

            var newValue = value;
            var isPercent = false;
            if (value.EndsWith(" %", StringComparison.InvariantCultureIgnoreCase))
            {
                isPercent = true;
                newValue = newValue.Replace(" %", "");
            }

            if (int.TryParse(newValue,
                             NumberStyles.AllowThousands | NumberStyles.AllowCurrencySymbol,
                             CultureInfo.InvariantCulture, out var intResult))
            {
                if (isPercent)
                {
                    return intResult / 100;
                }

                return intResult;
            }

            if (decimal.TryParse(newValue, NumberStyles.Number | NumberStyles.AllowCurrencySymbol,
                                 CultureInfo.InvariantCulture, out var decimalResult))
            {
                if (isPercent)
                {
                    return intResult / 100;
                }

                return decimalResult;
            }

            return value;
        }
    }
}