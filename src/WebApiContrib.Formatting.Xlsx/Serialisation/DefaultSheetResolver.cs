using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using WebApiContrib.Formatting.Xlsx.Attributes;

namespace WebApiContrib.Formatting.Xlsx.Serialisation
{
    public class DefaultSheetResolver : ISheetResolver
    {
        public ExcelSheetInfoCollection GetExcelSheetInfo(Type itemType, IEnumerable<object> data)
        {
            var sheets = FormatterUtils.GetMemberNames(itemType);
            var properties = GetSerialisablePropertyInfo(itemType, data);

            var sheetCollection = new  ExcelSheetInfoCollection();

            foreach (var sheet in sheets)
            {
                var prop = properties.FirstOrDefault(p => p.Name == sheet);

                if (prop == null) continue;

                sheetCollection.Add(new ExcelSheetInfo()
                {
                    SheetName = sheet,
                    ExcelSheetAttribute = FormatterUtils.GetAttribute<ExcelSheetAttribute>(prop)
                });
            }

            return sheetCollection;
        }

        public virtual IEnumerable<PropertyInfo> GetSerialisablePropertyInfo(Type itemType, IEnumerable<object> data)
        {
            return (from p in itemType.GetProperties()
                    where p.CanRead & p.GetGetMethod().IsPublic & p.GetGetMethod().GetParameters().Length == 0
                    select p).ToList();
        }
    }
}
