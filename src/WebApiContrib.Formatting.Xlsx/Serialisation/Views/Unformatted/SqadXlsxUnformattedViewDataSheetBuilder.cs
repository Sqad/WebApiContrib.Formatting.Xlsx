using System;
using System.Data;
using System.Xml;
using OfficeOpenXml;
using SQAD.MTNext.Serialisation.WebApiContrib.Formatting.Xlsx.Serialisation;
using SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Base;

namespace SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Views.Unformatted
{
    public class SqadXlsxUnformattedViewDataSheetBuilder : SqadXlsxSheetBuilderBase
    {
        private readonly string _dataUrl;

        public SqadXlsxUnformattedViewDataSheetBuilder(string dataUrl)
            : base(ExportViewConstants.UnformattedViewDataSheetName, shouldAutoFit: false)
        {
            _dataUrl = dataUrl;
        }

        protected override void CompileSheet(ExcelWorksheet worksheet, DataTable table)
        {
            if (table.Rows.Count == 0)
            {
                return;
            }

            WorksheetDataHelper.FillData(worksheet, table, true);

            FormatDateTimeColumns(worksheet);
        }
        
        private void FormatDateTimeColumns(ExcelWorksheet worksheet)
        {
            for (var i = 0; i < CurrentTable.Columns.Count; i++)
            {
                if (!(((ExcelCell) CurrentTable.Rows[0][i]).CellValue is DateTime))
                {
                    continue;
                }

                var column = worksheet.Cells[2, i + 1, worksheet.Dimension.Rows, i + 1];
                column.Style.Numberformat.Format = "mm-dd-yy";
            }
        }
    }
}
