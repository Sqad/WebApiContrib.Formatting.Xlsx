using SQAD.MTNext.Serialisation.WebApiContrib.Formatting.Xlsx.Serialisation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SQAD.MTNext.Interfaces.WebApiContrib.Formatting.Xlsx.Interfaces
{
    public interface ISheetResolver
    {
        ExcelSheetInfoCollection GetExcelSheetInfo(Type itemType, object data);
    }
}
