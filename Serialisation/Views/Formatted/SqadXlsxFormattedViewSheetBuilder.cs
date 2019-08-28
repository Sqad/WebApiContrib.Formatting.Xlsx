using System.Data;
using System.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Base;
using SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Views.Helpers;

namespace SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Views.Formatted
{
    public sealed class SqadXlsxFormattedViewSheetBuilder : SqadXlsxSheetBuilderBase
    {
        private readonly int _headerRowsCount;
        private int _leftPaneWidth;

        public SqadXlsxFormattedViewSheetBuilder(int headerRowsCount)
            : base(ExportViewConstants.FormattedViewSheetName)
        {
            _headerRowsCount = headerRowsCount;
        }

        protected override void CompileSheet(ExcelWorksheet worksheet, DataTable table)
        {
            if (table.Rows.Count == 0)
            {
                return;
            }

            WorksheetDataHelper.FillData(worksheet, table, false);

            ObtainLeftPaneWidth(worksheet);

            FormatHeader(worksheet);
            FormatRows(worksheet);
        }

        protected override void PostCompileActions(ExcelWorksheet worksheet)
        {
            worksheet.View.FreezePanes(_headerRowsCount + 1, _leftPaneWidth + 1);
        }

        private void ObtainLeftPaneWidth(ExcelWorksheet worksheet)
        {
            _leftPaneWidth = 3;
            for (var columnIndex = 0; columnIndex < worksheet.Dimension.Columns; columnIndex++)
            {
                var cell = worksheet.Cells[_headerRowsCount, columnIndex + 1];
                if (cell.Value == null)
                {
                    continue;
                }

                _leftPaneWidth = columnIndex;
                break;
            }
        }

        private void FormatRows(ExcelWorksheet sheet)
        {
            var firstDataRowIndex = _headerRowsCount + 1;
            var allDataCells = sheet.Cells[firstDataRowIndex, 1, sheet.Dimension.Rows, sheet.Dimension.Columns];
            allDataCells.Style.Fill.PatternType = ExcelFillStyle.Solid;
            allDataCells.Style.Fill.BackgroundColor.SetColor(Color.White);
            allDataCells.Style.Border.BorderAround(ExcelBorderStyle.None);

            var leftPaneCells = sheet.Cells[firstDataRowIndex, 1, sheet.Dimension.Rows, _leftPaneWidth];
            leftPaneCells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

            var percentsCells = sheet.Cells[firstDataRowIndex, 1, sheet.Dimension.Rows, 1];
            percentsCells.Style.Fill.PatternType = ExcelFillStyle.Solid;
            percentsCells.Style.Fill.BackgroundColor.SetColor(WorksheetHelpers.DataTotalsBackgroundColor);

            var beginningCells = sheet.Cells[firstDataRowIndex, _leftPaneWidth, sheet.Dimension.Rows, _leftPaneWidth];
            FormatNumbers(beginningCells);

            if (sheet.Dimension.Columns > _leftPaneWidth)
            {
                var dataCells = sheet.Cells[firstDataRowIndex, _leftPaneWidth + 1, sheet.Dimension.Rows,
                                            sheet.Dimension.Columns];
                FormatNumbers(dataCells);
            }

            FormatTotalColumns(sheet, firstDataRowIndex);
            FormatRows(sheet, firstDataRowIndex);
        }

        private void FormatTotalColumns(ExcelWorksheet sheet, int firstDataRowIndex)
        {
            for (var cellIndex = _leftPaneWidth + 1; cellIndex <= sheet.Dimension.Columns; cellIndex++)
            {
                var headerCell = sheet.Cells[_headerRowsCount, cellIndex];
                if (!headerCell.Merge)
                {
                    continue;
                }

                var totalCells = sheet.Cells[firstDataRowIndex, cellIndex, sheet.Dimension.Rows, cellIndex];
                totalCells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                totalCells.Style.Fill.BackgroundColor.SetColor(WorksheetHelpers.DataTotalsBackgroundColor);
            }
        }

        private void FormatHeader(ExcelWorksheet worksheet)
        {
            var allHeaderCells = worksheet.Cells[1, 1, _headerRowsCount, worksheet.Dimension.Columns];
            allHeaderCells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            allHeaderCells.Style.VerticalAlignment = ExcelVerticalAlignment.Distributed;

            allHeaderCells.Style.Fill.PatternType = ExcelFillStyle.Solid;
            allHeaderCells.Style.Fill.BackgroundColor.SetColor(WorksheetHelpers.HeaderBackgroundColor);

            allHeaderCells.Style.Font.Color.SetColor(WorksheetHelpers.HeaderFontColor);

            var mergedCells = worksheet.Cells[1, 1, _headerRowsCount, _leftPaneWidth];
            mergedCells.Merge = true;
            SetBordersToCells(mergedCells);

            for (var rowIndex = 1; rowIndex <= _headerRowsCount - 1; rowIndex++)
            {
                var startColumnIndex = _leftPaneWidth + 1;
                for (var endColumnIndex = startColumnIndex;
                     endColumnIndex <= worksheet.Dimension.Columns;
                     endColumnIndex++)
                {
                    var initialCell = worksheet.Cells[rowIndex, startColumnIndex];
                    if (initialCell.Value != null)
                    {
                        if (!(initialCell.Value is string)
                            || !((string) initialCell.Value).Contains(WorksheetHelpers.TotalRowIndicator)
                            || initialCell.Merge)
                        {
                            startColumnIndex++;
                            endColumnIndex = startColumnIndex - 1;

                            continue;
                        }

                        var verticalCellsToMerge =
                            worksheet.Cells[rowIndex, endColumnIndex, _headerRowsCount, endColumnIndex];
                        verticalCellsToMerge.Value = worksheet.Cells[rowIndex, endColumnIndex].Value;
                        verticalCellsToMerge.Merge = true;
                        verticalCellsToMerge.Style.Fill.BackgroundColor.SetColor(WorksheetHelpers
                                                                                     .HeaderTotalsBackgroundColor);
                        SetBordersToCells(verticalCellsToMerge);

                        startColumnIndex++;
                        endColumnIndex = startColumnIndex - 1;

                        continue;
                    }

                    var endCell = worksheet.Cells[rowIndex, endColumnIndex];
                    if (endCell.Value == null)
                    {
                        continue;
                    }

                    if (startColumnIndex != endColumnIndex)
                    {
                        var horizontalCellsToMerge =
                            worksheet.Cells[rowIndex, startColumnIndex, rowIndex, endColumnIndex];
                        horizontalCellsToMerge.Value = endCell.Value;
                        horizontalCellsToMerge.Merge = true;
                        SetBordersToCells(horizontalCellsToMerge);
                    }

                    startColumnIndex = endColumnIndex + 1;
                }
            }

            for (var columnIndex = _leftPaneWidth + 1; columnIndex <= worksheet.Dimension.Columns; columnIndex++)
            {
                SetBordersToCells(worksheet.Cells[_headerRowsCount, columnIndex]);
            }
        }

        private void FormatRows(ExcelWorksheet sheet, int firstDataRowIndex)
        {
            var isPreviousRowTotal = false;
            for (var rowIndex = firstDataRowIndex; rowIndex <= sheet.Dimension.Rows; rowIndex++)
            {
                var row = sheet.Cells[rowIndex, 1, rowIndex, sheet.Dimension.Columns];

                var nameCell = sheet.Cells[rowIndex, WorksheetHelpers.RowNameColumnIndex];
                var valueCell = sheet.Cells[rowIndex, _leftPaneWidth + 1];
                if (nameCell.Value != null
                    && nameCell.Value is string name
                    && name.Contains(WorksheetHelpers.TotalRowIndicator)
                    && valueCell.Value != null)
                {
                    FormatTotalRow(row);
                    isPreviousRowTotal = true;
                    continue;
                }

                if (isPreviousRowTotal)
                {
                    if (nameCell.Value == null)
                    {
                        FormatTotalRow(row);
                        continue;
                    }

                    isPreviousRowTotal = false;
                }

                var isGroupingRow = true;
                for (var x = WorksheetHelpers.RowNameColumnIndex + 1; x <= sheet.Dimension.Columns; x++)
                {
                    if (sheet.Cells[rowIndex, x].Value == null)
                    {
                        continue;
                    }

                    isGroupingRow = false;
                    break;
                }

                if (isGroupingRow)
                {
                    FormatGroupRow(row);
                }
            }
        }

        private void FormatTotalRow(ExcelRange row)
        {
            row.Style.Fill.PatternType = ExcelFillStyle.Solid;
            row.Style.Fill.BackgroundColor.SetColor(WorksheetHelpers.TotalsBackgroundColor);
        }

        private static void FormatGroupRow(ExcelRange row)
        {
            row.Style.Font.Bold = true;
        }

        private static void FormatNumbers(ExcelRange cells)
        {
            cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
        }

        private static void SetBordersToCells(ExcelRange cells)
        {
            cells.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.White);
        }
    }
}