using System.Data;
using System.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using OfficeOpenXml.Table.PivotTable;
using SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Base;

namespace SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Views.Unformatted
{
    public class SqadXlsxUnformattedViewPivotSheetBuilder : SqadXlsxSheetBuilderBase
    {
        public SqadXlsxUnformattedViewPivotSheetBuilder()
            : base(ExportViewConstants.UnformattedViewPivotSheetName, shouldAutoFit: false)
        {
        }

        protected override void CompileSheet(ExcelWorksheet worksheet, DataTable table)
        {
            var dataWorksheet = worksheet.Workbook
                                         .Worksheets
                                         .FirstOrDefault(x => x.Name == ExportViewConstants
                                                                  .UnformattedViewDataSheetName);
            if (dataWorksheet == null)
            {
                return;
            }

            var dataRange = dataWorksheet.Cells[dataWorksheet.Dimension.Address];

            worksheet.Cells[3, 1].Value = string.Empty;
            var pivotTable = worksheet.PivotTables.Add(worksheet.Cells[3, 1], dataRange, "PivotTable");

            pivotTable.StyleName = "PivotStyleLight16";
            pivotTable.TableStyle = TableStyles.Light16;
            //pivotTable.ShowHeaders = true;
            pivotTable.GridDropZones = false;
        }
    }
}
