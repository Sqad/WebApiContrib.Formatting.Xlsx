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
        //private  { get; set; }

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
            ExcelPackage package = Compile();

            return Task.Factory.StartNew(() => package.SaveAs(_stream));
        }

        private ExcelPackage Compile()
        {
            ExcelPackage package = new ExcelPackage();

            foreach (var sheet in _sheets.OrderBy(o=>o.IsReferenceSheet))
            {
                sheet.CompileSheet(package);
            }

            return package;
        }

        public bool IsExcelSupportedType(object expression)
        {
            return FormatterUtils.IsExcelSupportedType(expression);
        }
    }
}
