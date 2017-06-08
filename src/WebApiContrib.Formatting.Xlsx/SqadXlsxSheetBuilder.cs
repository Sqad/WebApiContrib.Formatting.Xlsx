using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WebApiContrib.Formatting.Xlsx
{
    public class SqadXlsxSheetBuilder
    {
        private ExcelWorksheet _worksheet { get; set; }
        private int _rowCount { get; set; }

        public SqadXlsxSheetBuilder(string SheetName)
        {
            _rowCount = 0;
            _worksheet = new ExcelWorksheet(null,null,null,null,SheetName,0,0,eWorkSheetHidden.Visible);
        }

        /// <summary>
        /// Append a row to the XLSX worksheet.
        /// </summary>
        /// <param name="row">The row to append to this instance.</param>
        public void AppendRow(IEnumerable<object> row)
        {
            _rowCount++;

            int i = 0;
            foreach (var col in row)
            {
                _worksheet.Cells[_rowCount, ++i].Value = col;
            }
        }

        public void FormatColumn(int column, string format, bool skipHeaderRow = true)
        {
            var firstRow = skipHeaderRow ? 2 : 1;

            if (firstRow <= _rowCount)
                _worksheet.Cells[firstRow, column, _rowCount, column].Style.Numberformat.Format = format;
        }
    }
}
