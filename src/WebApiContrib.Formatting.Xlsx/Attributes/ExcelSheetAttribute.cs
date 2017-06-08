using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WebApiContrib.Formatting.Xlsx.Attributes
{
    public class ExcelSheetAttribute : Attribute
    {
        private int? _order;

        public int Order
        {
            get { return _order ?? default(int); }
            set { _order = value; }
        }

        public ExcelSheetAttribute() { }

        public ExcelSheetAttribute(string SheetName) : this()
        {
            this.SheetName = SheetName;
        }

        public string SheetName { get; set; }
    }
}
