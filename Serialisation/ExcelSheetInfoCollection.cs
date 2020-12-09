
using SQAD.XlsxExportImport.Base.Models;

namespace SQAD.MTNext.Serialisation.WebApiContrib.Formatting.Xlsx.Serialisation
{
    public class ExcelSheetInfoCollection : KeyedCollectionBase<ExcelSheetInfo>
    {
        protected override string GetKeyForItem(ExcelSheetInfo item)
        {
            return item.PropertyName;
        }
    }
}
