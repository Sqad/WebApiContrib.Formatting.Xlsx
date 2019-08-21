using OfficeOpenXml;
using SQAD.MTNext.Serialisation.WebApiContrib.Formatting.Xlsx.Serialisation;
using SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Base;
using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace WebApiContrib.Formatting.Xlsx.src.WebApiContrib.Formatting.Xlsx.Serialisation.ApprovalReports
{
    public class SqadXlsxApprovalReportSheetBuilder : SqadXlsxSheetBuilderBase
    {
        private readonly int _startHeaderIndex;
        private readonly int _startDataIndex;
        private readonly int _totalCountColumns;
        private readonly int _totalCountRows;
        private readonly DateTime _startDateApprovalReport;
        private readonly DateTime _endDateApprovalReport;
        private readonly string _approvalType;

        public SqadXlsxApprovalReportSheetBuilder(int startHeaderIndex, int startDataIndex,
            int totalCountColumns, int totalCountRows, DateTime startDateApprovalReport,
            DateTime endDateApprovalReport, string approvalType) : base(ExportConstants.ApprovalReportSheetName)
        {
            _startHeaderIndex = startHeaderIndex;
            _startDataIndex = startDataIndex;
            _totalCountColumns = totalCountColumns;
            _totalCountRows = totalCountRows;
            _startDateApprovalReport = startDateApprovalReport;
            _endDateApprovalReport = endDateApprovalReport;
            _approvalType = approvalType;
        }

        private void FormatHeaderTemplate(ExcelWorksheet worksheet)
        {
            //Base configuration of excel document
            worksheet.SetValue(1, 1, "Appoval Report");
            worksheet.Cells[1, 1].Style.Font.Bold = true;

            worksheet.SetValue(2, 1, $"Date Range: {_startDateApprovalReport.ToString("MM/dd/yyyy")} to {_endDateApprovalReport.ToString("MM/dd/yyyy")}");
            //worksheet.Cells[1, 1].Style.Numberformat

            worksheet.SetValue(3, 1, $"Approval Type: {_approvalType}");
        }

        protected override void CompileSheet(ExcelWorksheet worksheet, DataTable table)
        {
            if (table.Rows.Count == 0)
            {
                return;
            }

            FormatHeaderTemplate(worksheet);

            for (int i = 0; i < _totalCountColumns; i++)
            {
                var column = table.Columns[i];
                var numberExcelColumn = i + 1;

                worksheet.SetValue(_startHeaderIndex, numberExcelColumn, column.ColumnName);
                worksheet.Cells[_startHeaderIndex,numberExcelColumn].Style.Font.Bold = true;

                for (var j = 0; j < _totalCountRows; j++)
                {
                    var row = table.Rows[j];
                    var rowValue = ((ExcelCell)row[column.ColumnName]).CellValue;
                    worksheet.SetValue(_startDataIndex + j, numberExcelColumn, rowValue);
                }
            }
        }
    }
}
