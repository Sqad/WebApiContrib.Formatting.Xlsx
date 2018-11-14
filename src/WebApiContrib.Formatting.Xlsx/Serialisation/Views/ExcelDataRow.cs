using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using SQAD.MTNext.Serialisation.WebApiContrib.Formatting.Xlsx.Serialisation;

namespace SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Views
{
    internal class ExcelDataRow
    {
        private readonly DataRow _dataRow;

        private string this[string columnName] => _dataRow.IsNull(columnName) ? null : (string) _dataRow[columnName];

        public ExcelDataRow(DataRow dataRow)
        {
            _dataRow = dataRow;
            IsHeader = bool.Parse(this["header"]);
        }

        public bool IsHeader { get; }

        public IEnumerable<ExcelCell> GetExcelCells(IEnumerable columns)
        {
            return from DataColumn column in columns
                select new ExcelCell
                       {
                           CellHeader = column.ColumnName,
                           CellValue = TryParseValue(this[column.ColumnName])
                       };
        }

        //note: DataTable values from ViewAPI already formatted, but Excel don't recognize it
        private static object TryParseValue(string value)
        {
            if (value == null)
            {
                return value;
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