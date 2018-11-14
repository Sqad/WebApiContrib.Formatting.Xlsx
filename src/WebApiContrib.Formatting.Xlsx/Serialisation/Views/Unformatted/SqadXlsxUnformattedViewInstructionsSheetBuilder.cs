using System.Data;
using OfficeOpenXml;
using SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Base;

namespace SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Views.Unformatted
{
    public class SqadXlsxUnformattedViewInstructionsSheetBuilder : SqadXlsxSheetBuilderBase
    {
        public SqadXlsxUnformattedViewInstructionsSheetBuilder(string sheetName)
            : base(sheetName)
        {
        }

        protected override void CompileSheet(ExcelWorksheet worksheet, DataTable table)
        {
            throw new System.NotImplementedException();
        }
    }
}
