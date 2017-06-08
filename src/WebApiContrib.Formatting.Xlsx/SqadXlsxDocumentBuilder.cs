using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WebApiContrib.Formatting.Xlsx
{
    public class SqadXlsxDocumentBuilder
    {
        private ExcelPackage Package { get; set; }
        
        private Stream _stream;

        public SqadXlsxDocumentBuilder(Stream stream)
        {
            _stream = stream;

            // Create a worksheet
            Package = new ExcelPackage();

        }

        public void AutoFit()
        {
            throw new NotImplementedException();
            //Worksheet.Cells[Worksheet.Dimension.Address].AutoFitColumns();
        }

        public Task WriteToStream()
        {
            return Task.Factory.StartNew(() => Package.SaveAs(_stream));
        }
        

        public static bool IsExcelSupportedType(object expression)
        {
            return expression is string
                || expression is short
                || expression is int
                || expression is long
                || expression is decimal
                || expression is float
                || expression is double
                || expression is DateTime;
        }
    }
}
