using OfficeOpenXml;
using OfficeOpenXml.Style;
using SQAD.MTNext.Business.Models.FlowChart.Enums;
using SQAD.MTNext.Serialisation.WebApiContrib.Formatting.Xlsx.Serialisation;
using SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Base;
using System;
using System.Data;
using SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.ApprovalReports.Helpers;
using SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.ApprovalReports.Enums;
using System.Drawing;

namespace SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.ApprovalReports
{
    public class SqadXlsxBillingReportSheetBuilder : SqadXlsxSheetBuilderBase
    {
        private readonly int _startHeaderIndex;
        private readonly int _startDataIndex;
        private readonly int _totalCountColumns;
        private readonly int _totalCountRows;

        public SqadXlsxBillingReportSheetBuilder(int startHeaderIndex, int startDataIndex,
            int totalCountColumns, int totalCountRows) : base(ExportConstants.BillingReportSheetName)
        {
            _startHeaderIndex = startHeaderIndex;
            _startDataIndex = startDataIndex;
            _totalCountColumns = totalCountColumns;
            _totalCountRows = totalCountRows;
        }
        private void FormatHeaderTemplate(ExcelWorksheet worksheet)
        {
            //Base configuration of excel document
            worksheet.SetValue(1, 1, "Billing Report");
            worksheet.Cells[1, 1].Style.Font.Bold = true;
        }

        private void FillWorksheetData(ExcelWorksheet worksheet, DataTable table)
        {
            for (int i = 0; i < _totalCountColumns; i++)
            {
                var column = table.Columns[i];
                var numberExcelColumn = i + 1;
                var headerCell = worksheet.Cells[_startHeaderIndex, numberExcelColumn];

                headerCell.Value = column.ColumnName;
                headerCell.Style.Font.Bold = true;

                for (var j = 0; j < _totalCountRows; j++)
                {
                    var row = table.Rows[j];
                    var rowValue = ((ExcelCell)row[column.ColumnName]).CellValue;
                    var rowCell = worksheet.Cells[_startDataIndex + j, numberExcelColumn];

                    rowCell.Value = rowValue;
                }
            }
        }

        protected override void CompileSheet(ExcelWorksheet worksheet, DataTable table)
        {
            if (table.Rows.Count == 0)
            {
                return;
            }

            FormatHeaderTemplate(worksheet);

            FillWorksheetData(worksheet, table);
        }
    }
}
