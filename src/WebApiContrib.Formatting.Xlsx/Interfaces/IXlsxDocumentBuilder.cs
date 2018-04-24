using OfficeOpenXml;
using SQAD.MTNext.WebApiContrib.Formatting.Xlsx;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SQAD.MTNext.Interfaces.WebApiContrib.Formatting.Xlsx.Interfaces
{
    public interface IXlsxDocumentBuilder
    {
        Task WriteToStream();

        bool IsExcelSupportedType(object expression);

        void AppendSheet(SqadXlsxSheetBuilder sheet);

        SqadXlsxSheetBuilder GetReferenceSheet();

        SqadXlsxSheetBuilder GetSheetByName(string name);
    }
}
