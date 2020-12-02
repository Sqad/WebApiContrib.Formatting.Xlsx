using SQAD.MTNext.Business.Models.Attributes;
using System;

namespace SQAD.MTNext.Serialisation.WebApiContrib.Formatting.Xlsx.Serialisation
{
    public class ExcelSheetInfo
    {
        public Type SheetType { get; set; }
        public object SheetObject { get; set; }
        public string PropertyName { get; set; }
        public ExcelSheetAttribute ExcelSheetAttribute { get; set; }
        public string SheetName { get; set; }
    }
}
