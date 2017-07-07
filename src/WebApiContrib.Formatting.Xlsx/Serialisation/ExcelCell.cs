using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WebApiContrib.Formatting.Xlsx.Serialisation
{
    public class ExcelCell
    {
        public string CellHeader { get; set; }
        public object CellValue { get; set; }
        public string DataValidationSheet { get; set; }
        public string DataValidationValue { get; set; } = "ID";
        public string DataValidationName { get; set; } = "Name";
    }
}
