using System;
using System.Data;
using OfficeOpenXml;
using SQAD.MTNext.Serialisation.WebApiContrib.Formatting.Xlsx.Serialisation;
using SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Base;

namespace SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Views.Unformatted
{
    public class SqadXlsxUnformattedViewDataSheetBuilder : SqadXlsxSheetBuilderBase
    {
        public SqadXlsxUnformattedViewDataSheetBuilder(string viewLabel = null)
            : base(string.IsNullOrEmpty(viewLabel) 
                   ? ExportViewConstants.UnformattedViewDataSheetName : viewLabel, shouldAutoFit: false)
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
