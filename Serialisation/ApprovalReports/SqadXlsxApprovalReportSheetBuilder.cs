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

            worksheet.SetValue(3, 1, $"Approval Type: {_approvalType}");
        }

        private void FormatNumber(ExcelRange cells, string format)
        {
            cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            cells.Style.Numberformat.Format = format;
        }

        private void FormatNumberCell(ExcelWorksheet worksheet, DataTable dataTable, int currentRowWorksheet, int currentRowDataTable, int currentColumnWorksheet,
                                      int currentColumnDataTable, ExcelFormatType excelFormatType)
        {
            var columnCount = dataTable.Columns.Count;

            var worksheetCell = worksheet.Cells[currentRowWorksheet, currentColumnWorksheet];
            var dataItem = currentColumnDataTable < columnCount ? dataTable.Rows[currentRowDataTable][currentColumnDataTable] : null;
            var excelCell = (ExcelCell)dataItem;

            if (excelCell != null)
            {
                switch (excelFormatType)
                {
                    case ExcelFormatType.IntNullable:
                        worksheetCell.Value = Parser.ParseNullableInt(excelCell.CellValue?.ToString());
                        FormatNumber(worksheetCell, ExportConstants.IntExcelFormatTemplate);
                        break;
                    case ExcelFormatType.Date:
                        worksheetCell.Value = DateTime.Parse(excelCell.CellValue.ToString());
                        FormatNumber(worksheetCell, ExportConstants.DateExcelFormatTemplate);
                        break;
                    case ExcelFormatType.DateNullable:
                        worksheetCell.Value = Parser.ParseNullableDateTime(excelCell.CellValue?.ToString());
                        FormatNumber(worksheetCell, ExportConstants.DateExcelFormatTemplate);
                        break;
                    case ExcelFormatType.AccountingNullable:
                        var currencySymbol = (ExcelCell)dataTable.Rows[currentRowDataTable][ExportConstants.CurrencySymbolColumnName];

                        worksheetCell.Value = Parser.ParseNullableFloat(excelCell.CellValue?.ToString());
                        FormatNumber(worksheetCell, ExportConstants.CreateAccountingExcelFormatTemplate(currencySymbol.CellValue?.ToString()));
                        break;
                    default: throw new Exception("Inccorect excel format type");
                }
            }
        }

        private void FormatNumberData(ExcelWorksheet worksheet, DataTable dataTable)
        {
            var startRowWorksheet = _startHeaderIndex + 1;

            //Datatable starts calculate index from 0 
            int numberDayColumnDataTable = (int)ApprovalReportElement.Days,
                numberDateSubmitedColumnDataTable = (int)ApprovalReportElement.DateSubmitted,
                numberDateCompletedColumnDataTable = (int)ApprovalReportElement.DateCompleted,
                numberGrossCostColumnDataTable = (int)ApprovalReportElement.GrossCost,
                //For delete one column gross or net cost we should use minus 1
                numberWorkingCostColumnDataTable = (int)ApprovalReportElement.WorkingCost - 1,
                numberNonWorkingCostColumnDataTable = (int)ApprovalReportElement.NonWorkingCosts - 1,
                numberFeesCostColumnDataTable = (int)ApprovalReportElement.Fees - 1;

            //Worksheet starts calculate index from 1 
            int numberDayColumnWorksheet = numberDayColumnDataTable + 1,
                numberDateSubmitedColumnWorksheet = numberDateSubmitedColumnDataTable + 1,
                numberDateCompletedColumnWorksheet = numberDateCompletedColumnDataTable + 1,
                numberGrossCostColumnWorksheet = numberGrossCostColumnDataTable + 1,
                numberWorkingCostColumnWorksheet = numberWorkingCostColumnDataTable + 1,
                numberNonWorkingCostColumnWorksheet = numberNonWorkingCostColumnDataTable + 1,
                numberFeesCostColumnWorksheet = numberFeesCostColumnDataTable + 1;

            for (int i = 0; i < dataTable.Rows.Count; i++)
            {
                //Days
                FormatNumberCell(worksheet: worksheet, dataTable: dataTable,
                                 currentRowWorksheet: startRowWorksheet + i, currentRowDataTable: i,
                                 currentColumnWorksheet: numberDayColumnWorksheet, currentColumnDataTable: numberDayColumnDataTable,
                                 excelFormatType: ExcelFormatType.IntNullable);

                //DateSubmitted
                FormatNumberCell(worksheet: worksheet, dataTable: dataTable,
                               currentRowWorksheet: startRowWorksheet + i, currentRowDataTable: i,
                               currentColumnWorksheet: numberDateSubmitedColumnWorksheet, currentColumnDataTable: numberDateSubmitedColumnDataTable,
                               excelFormatType: ExcelFormatType.Date);

                //DateCompleted
                FormatNumberCell(worksheet: worksheet, dataTable: dataTable,
                             currentRowWorksheet: startRowWorksheet + i, currentRowDataTable: i,
                             currentColumnWorksheet: numberDateCompletedColumnWorksheet, currentColumnDataTable: numberDateCompletedColumnDataTable,
                             excelFormatType: ExcelFormatType.DateNullable);
                //Gross or Net cost
                FormatNumberCell(worksheet: worksheet, dataTable: dataTable,
                           currentRowWorksheet: startRowWorksheet + i, currentRowDataTable: i,
                           currentColumnWorksheet: numberGrossCostColumnWorksheet, currentColumnDataTable: numberGrossCostColumnDataTable,
                           excelFormatType: ExcelFormatType.AccountingNullable);

                //Working cost
                FormatNumberCell(worksheet: worksheet, dataTable: dataTable,
                           currentRowWorksheet: startRowWorksheet + i, currentRowDataTable: i,
                           currentColumnWorksheet: numberWorkingCostColumnWorksheet, currentColumnDataTable: numberWorkingCostColumnDataTable,
                           excelFormatType: ExcelFormatType.AccountingNullable);

                //Non-Working cost
                FormatNumberCell(worksheet: worksheet, dataTable: dataTable,
                           currentRowWorksheet: startRowWorksheet + i, currentRowDataTable: i,
                           currentColumnWorksheet: numberNonWorkingCostColumnWorksheet, currentColumnDataTable: numberNonWorkingCostColumnDataTable,
                           excelFormatType: ExcelFormatType.AccountingNullable);

                //Fees cost
                FormatNumberCell(worksheet: worksheet, dataTable: dataTable,
                           currentRowWorksheet: startRowWorksheet + i, currentRowDataTable: i,
                           currentColumnWorksheet: numberFeesCostColumnWorksheet, currentColumnDataTable: numberFeesCostColumnDataTable,
                           excelFormatType: ExcelFormatType.AccountingNullable);
            }
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
                    if ((bool.Parse(((ExcelCell)row[ExportConstants.EvenGroupColumnName]).CellValue.ToString())) == true)
                    {
                        rowCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        rowCell.Style.Fill.BackgroundColor.SetColor(ExportConstants.EvenGroupDefaultColor);
                    }
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

            FormatNumberData(worksheet, table);
        }
    }
}
