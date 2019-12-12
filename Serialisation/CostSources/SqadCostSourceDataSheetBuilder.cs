using System.Data;
using System.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Base;

namespace SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.CostSources
{
    internal class SqadCostSourceDataSheetBuilder : SqadXlsxSheetBuilderBase
    {
        private readonly DataTable _dataTable;
        private readonly int _periodsCount;

        public SqadCostSourceDataSheetBuilder(string sheetName,
                                              DataTable dataTable,
                                              int periodsCount,
                                              bool isReferenceSheet = false,
                                              bool isPreservationSheet = false,
                                              bool isHidden = false,
                                              bool shouldAutoFit = true)
            : base(sheetName, isReferenceSheet, isPreservationSheet, isHidden, shouldAutoFit)
        {
            this._dataTable = dataTable;
            _periodsCount = periodsCount;
        }

        protected override void CompileSheet(ExcelWorksheet worksheet, DataTable table)
        {
            worksheet.Cells.LoadFromDataTable(_dataTable, true);

            worksheet.Cells.Style.Font.Name = "Calibri";

            var headerRow = worksheet.Row(1);
            headerRow.Style.Fill.PatternType = ExcelFillStyle.Solid;
            headerRow.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(122, 122, 122));
            headerRow.Style.Font.Color.SetColor(Color.White);

            if (table.Rows.Count < 1)
            {
                return;
            }

            var periodCells = worksheet.Cells[2, 
                                              worksheet.Dimension.Columns - _periodsCount,
                                              worksheet.Dimension.Rows,
                                              worksheet.Dimension.Columns];
            periodCells.Style.Numberformat.Format = "0.00";
        }
    }
}