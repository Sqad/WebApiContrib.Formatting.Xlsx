using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SQAD.MTNext.Serialisation.WebApiContrib.Formatting.Xlsx.Serialisation
{
    public class ExcelCell
    {
        public string CellHeader { get; set; }
        public object CellValue { get; set; }
        public string DataValidationSheet { get; set; }
        public int DataValidationValueCellIndex { get; set; }
        public int DataValidationNameCellIndex { get; set; }
        public int DataValidationBeginRow { get; set; }
        public int DataValidationRowsCount { get; set; }
    }
}
