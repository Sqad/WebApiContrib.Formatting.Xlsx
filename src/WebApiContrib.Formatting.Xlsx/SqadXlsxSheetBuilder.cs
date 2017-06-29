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
        private string _sheetName { get; set; }
        public string SheetName => _sheetName;

        private  bool _isReferenceSheet { get; set; }
        public bool IsReferenceSheet => _isReferenceSheet;

        public bool ShouldAutoFit { get; set; }

        private List<Dictionary<int, object>> _valueByColumnNumber { get; set; }

        public SqadXlsxSheetBuilder(string SheetName, bool IsReferenceSheet=false)
        {
            _sheetName = SheetName;
            _isReferenceSheet = IsReferenceSheet;
            _valueByColumnNumber = new List<Dictionary<int, object>>();
        }

        /// <summary>
        /// Append a row to the XLSX worksheet.
        /// </summary>
        /// <param name="row">The row to append to this instance.</param>
        public void AppendRow(IEnumerable<object> row)
        {
            Dictionary<int, object> newRow = new Dictionary<int, object>();

            foreach (var colValue in row)
            {
                //new ExcelWorksheet().Cells[]
                //_worksheet.Cells[_rowCount, ++i].Value = col;
                newRow.Add(newRow.Count + 1, colValue);
            }

            _valueByColumnNumber.Add(newRow);
        }

        public void CompileSheet(ExcelPackage package)
        {
            if (_valueByColumnNumber.Count() == 0)
                return ;

            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(SheetName) ;

            int rowCount = 0;
            foreach (var row in _valueByColumnNumber)
            {
                rowCount++;
                foreach (var col in row)
                {
                    worksheet.Cells[rowCount, col.Key].Value = col.Value;
                }
            }

            if (worksheet.Dimension != null && ShouldAutoFit)
                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

        }

        //public void FormatColumn(int column, string format, bool skipHeaderRow = true)
        //{
        //    var firstRow = skipHeaderRow ? 2 : 1;

        //    if (firstRow <= _rowCount)
        //        _worksheet.Cells[firstRow, column, _rowCount, column].Style.Numberformat.Format = format;
        //}

    }
}
