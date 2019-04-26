using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using SQAD.MTNext.Serialisation.WebApiContrib.Formatting.Xlsx.Serialisation;

namespace SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Views
{
    public class ExcelDataRow
    {
        private readonly DataRow _dataRow;

        protected object this[string columnName] => _dataRow.IsNull(columnName) ? null : _dataRow[columnName];

        public ExcelDataRow(DataRow dataRow)
        {
            _dataRow = dataRow;
        }

        public IEnumerable<ExcelCell> GetExcelCells(IEnumerable columns)
        {
            return from DataColumn column in columns
                select new ExcelCell
                       {
                           CellHeader = column.ColumnName,
                           CellValue = ParseValue(column.ColumnName)
                       };
        }

        protected virtual object ParseValue(string columnName)
        {
            return this[columnName];
        }
    }
}
