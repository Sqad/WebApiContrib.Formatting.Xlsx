using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WebApiContrib.Formatting.Xlsx.Serialisation
{
    public interface ISheetResolver
    {
        ExcelSheetInfoCollection GetExcelSheetInfo(Type itemType, object data);
    }
}
