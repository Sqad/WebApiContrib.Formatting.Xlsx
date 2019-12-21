using System.Threading.Tasks;
using SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Base;
using WebApiContrib.Formatting.Xlsx.Models;

namespace SQAD.MTNext.Interfaces.WebApiContrib.Formatting.Xlsx.Interfaces
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
