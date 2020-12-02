using SQAD.MTNext.Serialisation.WebApiContrib.Formatting.Xlsx.Serialisation;
using System;

namespace SQAD.MTNext.Interfaces.WebApiContrib.Formatting.Xlsx.Interfaces
{
    public interface ISheetResolver
    {
        ExcelSheetInfoCollection GetExcelSheetInfo(Type itemType, object data);
    }
}
