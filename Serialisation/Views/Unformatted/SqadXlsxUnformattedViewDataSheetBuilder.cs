using OfficeOpenXml;
using SQAD.MTNext.Serialisation.WebApiContrib.Formatting.Xlsx.Serialisation;
using SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Base;
using System;
using System.Data;

namespace SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Views.Unformatted
{
    public class SqadXlsxUnformattedViewDataSheetBuilder : SqadXlsxSheetBuilderBase
    {
        public SqadXlsxUnformattedViewDataSheetBuilder()
            : base(ExportViewConstants.UnformattedViewDataSheetName, shouldAutoFit: false)
        {
        }

        protected override void CompileSheet(ExcelWorksheet worksheet, DataTable table)
        {
            if (table.Rows.Count == 0)
            {
                return;
            }

            WorksheetDataHelper.FillData(worksheet, table, true);

            FormatColumns(worksheet);
        }
        
        private void FormatColumns(ExcelWorksheet worksheet)
        {
            for (var i = 0; i < CurrentTable.Columns.Count; i++)
            {
                var columnValue = ((ExcelCell) CurrentTable.Rows[0][i]).CellValue;
                if (columnValue is double)
                {
                    //note: force EPPlus to DON'T ROUND NUMBERS
                    worksheet.Cells[2, i + 1, worksheet.Dimension.Rows, i + 1].Style.Numberformat.Format = "0.00";
                    continue;
                }

                if (!(columnValue is DateTime))
                {
                    continue;
                }

                var column = worksheet.Cells[2, i + 1, worksheet.Dimension.Rows, i + 1];
                column.Style.Numberformat.Format = "mm-dd-yy";
            }
        }
    }
}
