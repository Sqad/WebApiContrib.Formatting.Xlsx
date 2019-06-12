using System.Data;
using OfficeOpenXml;
using SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Base;

namespace SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Views.Unformatted
{
    public class SqadXlsxUnformattedViewInstructionsSheetBuilder : SqadXlsxSheetBuilderBase
    {
        public SqadXlsxUnformattedViewInstructionsSheetBuilder()
            : base(ExportViewConstants.UnformattedViewInstructionsSheetName)
        {
        }

        protected override void CompileSheet(ExcelWorksheet worksheet, DataTable table)
        {
            WorksheetDataHelper.FillData(worksheet, table, false);

            var headerRow = worksheet.Cells[1, 1, 1, worksheet.Dimension.Columns];
            headerRow.Style.Font.Size = 20;
            headerRow.Style.Font.Bold = true;

            var subHeaderRow = worksheet.Cells[2, 1, 2, worksheet.Dimension.Columns];
            subHeaderRow.Style.Font.Size = 14;
            subHeaderRow.Style.Font.Bold = true;

            worksheet.Column(1).Width = 20;
        }
    }
}
