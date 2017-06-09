using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WebApiContrib.Formatting.Xlsx.Serialisation
{
    public class ExcelSheetInfoCollection : KeyedCollectionBase<ExcelSheetInfo>
    {
        protected override string GetKeyForItem(ExcelSheetInfo item)
        {
            return item.PropertyName;
        }
    }
}
