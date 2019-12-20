using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using OfficeOpenXml;
using SQAD.MTNext.Interfaces.WebApiContrib.Formatting.Xlsx.Interfaces;
using SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Base;
using WebApiContrib.Formatting.Xlsx.Models;

namespace SQAD.MTNext.WebApiContrib.Formatting.Xlsx
{
    public class SqadXlsxDocumentBuilder : IXlsxDocumentBuilder
    {
        //private  { get; set; }

        private Stream _stream;


        //New Stuff
        private List<SqadXlsxSheetBuilderBase> _sheets { get; set; }

        private XlsxTemplateInfo _templateInfo;

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

        public SqadXlsxSheetBuilderBase GetPreservationSheet() => _sheets.FirstOrDefault(w => w.IsPreservationSheet);

        public SqadXlsxSheetBuilderBase GetSheetByName(string name)
        {
            return _sheets.FirstOrDefault(w => w.ContainsTable(name));
        }

        public bool IsVBA => _sheets.Any(a => a.IsReferenceSheet);

        public void SetTemplateInfo(XlsxTemplateInfo templateInfo)
        {
            _templateInfo = templateInfo;
        }

        public async  Task WriteToStream()
        {
            var package = await Task.Run(() =>
                                         {
                                             Stream templateStream = null;
                                             ExcelPackage excelPackage;
                                             try
                                             {
                                                 if (_templateInfo == null)
                                                 {
                                                     excelPackage = new ExcelPackage();
                                                 }
                                                 else
                                                 {
                                                     templateStream = File.OpenRead(_templateInfo.Path);
                                                     excelPackage = new ExcelPackage(templateStream);
                                                 }
                                             }
                                             finally
                                             {
                                                 templateStream?.Dispose();
                                             }

                                             foreach (var sheet in _sheets.OrderBy(o => o.IsReferenceSheet))
                                             {
                                                 sheet.CompileSheet(excelPackage);
                                             }

                                             excelPackage.SaveAs(_stream);

                                             return excelPackage;
                                         });
        }

        public bool IsExcelSupportedType(object expression)
        {
            return FormatterUtils.IsExcelSupportedType(expression);
        }
    }
}