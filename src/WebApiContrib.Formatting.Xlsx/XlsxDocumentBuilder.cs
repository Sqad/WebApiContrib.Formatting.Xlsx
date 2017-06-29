using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using WebApiContrib.Formatting.Xlsx.Interfaces;
using WebApiContrib.Formatting.Xlsx.Serialisation;

namespace WebApiContrib.Formatting.Xlsx
{
    public class XlsxDocumentBuilder : IXlsxDocumentBuilder
    {
        public ExcelPackage Package { get; set; }
        public ExcelWorksheet Worksheet { get; set; }
        public int RowCount { get; set; }

        private Stream _stream;

        public XlsxDocumentBuilder(Stream stream)
        {
            _stream = stream;

            // Create a worksheet
            Package = new ExcelPackage();
            Package.Workbook.Worksheets.Add("Data");
            Worksheet = Package.Workbook.Worksheets[1];

            RowCount = 0;
        }

        public void AutoFit()
        {
            Worksheet.Cells[Worksheet.Dimension.Address].AutoFitColumns();
        }

        public Task WriteToStream()
        {
            return Task.Factory.StartNew(() => Package.SaveAs(_stream));
        }

        /// <summary>
        /// Append a row to the XLSX worksheet.
        /// </summary>
        /// <param name="row">The row to append to this instance.</param>
        public void AppendRow(IEnumerable<object> row)
        {
            RowCount++;
            
            int i = 0;
            foreach (var col in row)
            {
                Worksheet.Cells[RowCount, ++i].Value = col;
            }
        }

        public void FormatColumn(int column, string format, bool skipHeaderRow = true)
        {
            var firstRow = skipHeaderRow ? 2 : 1;

            if (firstRow <= RowCount)
                Worksheet.Cells[firstRow, column, RowCount, column].Style.Numberformat.Format = format;
        }

        public bool IsExcelSupportedType(object expression)
        {
            return FormatterUtils.IsExcelSupportedType(expression);
        }

        public void AppendSheet(SqadXlsxSheetBuilder sheet)
        {
            throw new NotImplementedException();
        }
    }
}
