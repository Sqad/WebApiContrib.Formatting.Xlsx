using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WebApiContrib.Formatting.Xlsx.Interfaces
{
    public interface IXlsxDocumentBuilder
    {
        Task WriteToStream();

        bool IsExcelSupportedType(object expression);

        void AppendSheet(SqadXlsxSheetBuilder sheet);

        SqadXlsxSheetBuilder GetReferenceSheet();
    }
}
