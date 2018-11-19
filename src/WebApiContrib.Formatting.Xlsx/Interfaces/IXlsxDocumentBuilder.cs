using System.Threading.Tasks;
using SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Base;

namespace SQAD.MTNext.Interfaces.WebApiContrib.Formatting.Xlsx.Interfaces
{
    public interface IXlsxDocumentBuilder
    {
        Task WriteToStream();

        bool IsExcelSupportedType(object expression);

        void AppendSheet(SqadXlsxSheetBuilderBase sheet);

        SqadXlsxSheetBuilderBase GetReferenceSheet();

        SqadXlsxSheetBuilderBase GetSheetByName(string name);
    }
}
