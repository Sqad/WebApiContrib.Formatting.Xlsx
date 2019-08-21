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
        public SqadXlsxApprovalReportSheetBuilder(int startHeaderIndex, int startDataIndex,
            int totalCountColumns, int totalCountRows, DateTime startDateApprovalReport,
            DateTime endDateApprovalReport) : base(ExportConstants.ApprovalReportSheetName)
        {
            _startHeaderIndex = startHeaderIndex;
            _startDataIndex = startDataIndex;
            _totalCountColumns = totalCountColumns;
            _totalCountRows = totalCountRows;
            _startDateApprovalReport = startDateApprovalReport;
            _endDateApprovalReport = endDateApprovalReport;
        }

        protected override void CompileSheet(ExcelWorksheet worksheet, DataTable table)
        {
            if (table.Rows.Count == 0)
            {
                return;
            }

            //Base configuration of excel document
            worksheet.SetValue(0, 0, "Appoval Report");
            worksheet.SetValue(1, 0, $"Date Range: {_startDateApprovalReport} to {_endDateApprovalReport}");


            for (int i = 0; i < _totalCountColumns; i++)
            {
                var column = table.Columns[i];
                var numberExcelColumn = i + 1;

                worksheet.SetValue(_startHeaderIndex, numberExcelColumn, column.ColumnName);

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
