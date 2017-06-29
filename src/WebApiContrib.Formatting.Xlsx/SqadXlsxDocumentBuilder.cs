using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WebApiContrib.Formatting.Xlsx.Interfaces;

namespace WebApiContrib.Formatting.Xlsx
{
    public class SqadXlsxDocumentBuilder : IXlsxDocumentBuilder
    {
        //private ExcelPackage Package { get; set; }

        private Stream _stream;


        //New Stuff
        private List<SqadXlsxSheetBuilder> _sheets { get; set; }



        public SqadXlsxDocumentBuilder(Stream stream)
        {
            _stream = stream;

            // Create a worksheet
            //Package = new ExcelPackage();

            _sheets = new List<SqadXlsxSheetBuilder>();
        }


        public void AppendSheet(SqadXlsxSheetBuilder sheet)
        {
            _sheets.Add(sheet);
        }


        public Task WriteToStream()
        {
            //return Task.Factory.StartNew(() => Package.SaveAs(_stream));
            return null;
        }

        //public ExcelWorksheet AppendSheet(string sheetName)
        //{
        //    //return Package.Workbook.Worksheets.Add(sheetName);
        //    return null;
        //}

        public bool IsExcelSupportedType(object expression)
        {
            return FormatterUtils.IsExcelSupportedType(expression);
        }
    }
}
