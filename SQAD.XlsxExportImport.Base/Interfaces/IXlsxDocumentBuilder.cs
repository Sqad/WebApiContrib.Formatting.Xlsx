using System.Threading.Tasks;
using SQAD.XlsxExportImport.Base.Builders;
using SQAD.XlsxExportImport.Base.Models;

namespace SQAD.XlsxExportImport.Base.Interfaces
{
    public interface IXlsxDocumentBuilder
    {
        void SetTemplateInfo(XlsxTemplateInfo templateInfo);

        Task WriteToStream();

        bool IsExcelSupportedType(object expression);

        void AppendSheet(SqadXlsxSheetBuilderBase sheet);

        SqadXlsxSheetBuilderBase GetReferenceSheet();

        SqadXlsxSheetBuilderBase GetPreservationSheet();

        SqadXlsxSheetBuilderBase GetSheetByName(string name);
    }
}
