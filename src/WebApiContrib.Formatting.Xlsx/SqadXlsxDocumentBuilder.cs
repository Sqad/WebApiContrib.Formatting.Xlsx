using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using OfficeOpenXml;
using SQAD.MTNext.Interfaces.WebApiContrib.Formatting.Xlsx.Interfaces;
using SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Base;

namespace SQAD.MTNext.WebApiContrib.Formatting.Xlsx
{
    public class SqadXlsxDocumentBuilder : IXlsxDocumentBuilder
    {
        //private  { get; set; }

        private Stream _stream;


        //New Stuff
        private List<SqadXlsxSheetBuilderBase> _sheets { get; set; }

        public SqadXlsxDocumentBuilder(Stream stream)
        {
            _stream = stream;
            _sheets = new List<SqadXlsxSheetBuilderBase>();
        }


        public void AppendSheet(SqadXlsxSheetBuilderBase sheet)
        {
            _sheets.Add(sheet);
        }

        public SqadXlsxSheetBuilderBase GetReferenceSheet() => _sheets.FirstOrDefault(w => w.IsReferenceSheet);

        public SqadXlsxSheetBuilderBase GetSheetByName(string name)
        {
            return _sheets.FirstOrDefault(w => w.ContainsTable(name));
        }

        public bool IsVBA => _sheets.Any(a => a.IsReferenceSheet);

        public Task WriteToStream()
        {
            ExcelPackage package = Compile();
            return Task.Factory.StartNew(() => package.SaveAs(_stream));
        }

        private ExcelPackage Compile()
        {
            ExcelPackage package = new ExcelPackage();

            foreach (var sheet in _sheets.OrderBy(o => o.IsReferenceSheet))
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