using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WebApiContrib.Formatting.Xlsx.Serialisation;

namespace WebApiContrib.Formatting.Xlsx.Interfaces
{
    public interface ISheetResolver
    {
        ExcelSheetInfoCollection GetExcelSheetInfo(Type itemType, object data);
    }
}
