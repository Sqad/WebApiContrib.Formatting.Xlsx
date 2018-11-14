﻿using System.Data;
using System.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using SQAD.MTNext.Serialisation.WebApiContrib.Formatting.Xlsx.Serialisation;
using SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Base;

namespace SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Views.Formatted
{
    public sealed class SqadXlsxFormattedViewSheetBuilder : SqadXlsxSheetBuilderBase
    {
        private const int LeftPaneWidth = 3;
        private const int RowNameColumnIndex = 2;
        private const string TotalRowIndicator = "Total";

        private readonly Color _headerBackgroundColor = Color.FromArgb(212, 227, 244);
        private readonly Color _headerTotalsBackgroundColor = Color.FromArgb(198, 217, 241);
        private readonly Color _headerFontColor = Color.FromArgb(47, 79, 79);
        private readonly Color _dataTotalsBackgroundColor = Color.FromArgb(244, 244, 244);
        private readonly Color _totalsBackgroundColor = Color.FromArgb(227, 236, 248);
        
        private readonly int _headerRowsCount;

        public SqadXlsxFormattedViewSheetBuilder(string sheetName, int headerRowsCount)
            : base(sheetName)
        {
            _headerRowsCount = headerRowsCount;
        }

        protected override void CompileSheet(ExcelWorksheet worksheet, DataTable table)
        {
            if (table.Rows.Count == 0)
            {
                return;
            }
            
            FillData(worksheet, table);

            FormatHeader(worksheet);
            FormatRows(worksheet);
        }

        protected override void PostCompileActions(ExcelWorksheet worksheet)
        {
            worksheet.View.FreezePanes(_headerRowsCount + 1, LeftPaneWidth + 1);
        }

        private void FormatRows(ExcelWorksheet sheet)
        {
            var firstDataRowIndex = _headerRowsCount + 1;
            var allDataCells = sheet.Cells[firstDataRowIndex, 1, sheet.Dimension.Rows, sheet.Dimension.Columns];
            allDataCells.Style.Fill.PatternType = ExcelFillStyle.Solid;
            allDataCells.Style.Fill.BackgroundColor.SetColor(Color.White);
            allDataCells.Style.Border.BorderAround(ExcelBorderStyle.None);

            var leftPaneCells = sheet.Cells[firstDataRowIndex, 1, sheet.Dimension.Rows, LeftPaneWidth];
            leftPaneCells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

            var percentsCells = sheet.Cells[firstDataRowIndex, 1, sheet.Dimension.Rows, 1];
            percentsCells.Style.Fill.PatternType = ExcelFillStyle.Solid;
            percentsCells.Style.Fill.BackgroundColor.SetColor(_dataTotalsBackgroundColor);
            percentsCells.Style.Numberformat.Format = "0 %";

            var dataCells = sheet.Cells[firstDataRowIndex, LeftPaneWidth + 1, sheet.Dimension.Rows,
                                        sheet.Dimension.Columns];
            dataCells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            dataCells.Style.Numberformat.Format = "#,###";
            
            FormatTotalColumns(sheet, firstDataRowIndex);
            FormatRows(sheet, firstDataRowIndex);
        }

        private void FormatTotalColumns(ExcelWorksheet sheet, int firstDataRowIndex)
        {
            for (var cellIndex = LeftPaneWidth + 1; cellIndex <= sheet.Dimension.Columns; cellIndex++)
            {
                var headerCell = sheet.Cells[_headerRowsCount, cellIndex];
                if (!headerCell.Merge)
                {
                    continue;
                }

                var totalCells = sheet.Cells[firstDataRowIndex, cellIndex, sheet.Dimension.Rows, cellIndex];
                totalCells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                totalCells.Style.Fill.BackgroundColor.SetColor(_dataTotalsBackgroundColor);
            }
        }

        private void FormatHeader(ExcelWorksheet worksheet)
        {
            var allHeaderCells = worksheet.Cells[1, 1, _headerRowsCount, worksheet.Dimension.Columns];
            allHeaderCells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            allHeaderCells.Style.VerticalAlignment = ExcelVerticalAlignment.Distributed;

            allHeaderCells.Style.Fill.PatternType = ExcelFillStyle.Solid;
            allHeaderCells.Style.Fill.BackgroundColor.SetColor(_headerBackgroundColor);

            allHeaderCells.Style.Font.Color.SetColor(_headerFontColor);

            var mergedCells = worksheet.Cells[1, 1, _headerRowsCount, LeftPaneWidth];
            mergedCells.Merge = true;
            SetBordersToCells(mergedCells);

            for (var rowIndex = 1; rowIndex <= _headerRowsCount - 1; rowIndex++)
            {
                var startColumnIndex = LeftPaneWidth + 1;
                for (var endColumnIndex = startColumnIndex;
                     endColumnIndex <= worksheet.Dimension.Columns;
                     endColumnIndex++)
                {
                    var initialCell = worksheet.Cells[rowIndex, startColumnIndex];
                    if (initialCell.Value != null)
                    {
                        startColumnIndex++;
                        endColumnIndex = startColumnIndex;

                        continue;
                    }

                    var endCell = worksheet.Cells[rowIndex, endColumnIndex];
                    if (endCell.Value == null)
                    {
                        continue;
                    }

                    var horizontalCellsToMerge = worksheet.Cells[rowIndex, startColumnIndex, rowIndex, endColumnIndex];
                    horizontalCellsToMerge.Value = endCell.Value;
                    horizontalCellsToMerge.Merge = true;
                    SetBordersToCells(horizontalCellsToMerge);

                    var verticalCellsToMerge =
                        worksheet.Cells[rowIndex, endColumnIndex + 1, _headerRowsCount, endColumnIndex + 1];
                    verticalCellsToMerge.Value = worksheet.Cells[rowIndex, endColumnIndex + 1].Value;
                    verticalCellsToMerge.Merge = true;
                    verticalCellsToMerge.Style.Fill.BackgroundColor.SetColor(_headerTotalsBackgroundColor);
                    SetBordersToCells(verticalCellsToMerge);

                    endColumnIndex += 2;
                    startColumnIndex = endColumnIndex;
                }
            }

            for (var columnIndex = LeftPaneWidth + 1; columnIndex <= worksheet.Dimension.Columns; columnIndex++)
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

                var nameCell = sheet.Cells[rowIndex, RowNameColumnIndex];
                if (nameCell.Value != null && nameCell.Value is string name && name.Contains(TotalRowIndicator))
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
                for (var x = RowNameColumnIndex + 1; x <= sheet.Dimension.Columns; x++)
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
            row.Style.Fill.BackgroundColor.SetColor(_totalsBackgroundColor);
        }

        private static void FormatGroupRow(ExcelRange row)
        {
            row.Style.Font.Bold = true;
        }

        private static void FillData(ExcelWorksheet worksheet, DataTable table)
        {
            for (var y = 0; y < table.Rows.Count; y++)
            {
                var dataRow = table.Rows[y];

                for (var x = 0; x < table.Columns.Count; x++)
                {
                    var column = table.Columns[x];
                    var value = (ExcelCell) dataRow[column.ColumnName];

                    var cell = worksheet.Cells[y + 1, x + 1];
                    cell.Value = value.CellValue;
                }
            }
        }

        private static void SetBordersToCells(ExcelRange cells)
        {
            cells.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.White);
        }
    }
}