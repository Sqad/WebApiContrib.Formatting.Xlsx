using System.Data;
using System.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using SQAD.MTNext.Serialisation.WebApiContrib.Formatting.Xlsx.Serialisation;
using SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Base;

namespace SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Views
{
    public sealed class SqadXlsxViewSheetBuilder : SqadXlsxSheetBuilderBase
    {
        private const int FirstColumnsToMerge = 3;

        private readonly int _headerRowsCount;

        private readonly Color _headerBackgroundColor = Color.FromArgb(212, 227, 244);
        private readonly Color _headerTotalsBackgroundColor = Color.FromArgb(198, 217, 241);
        private readonly Color _headerFontColor = Color.FromArgb(47, 79, 79);
        private readonly Color _dataTotalsBackgroundColor = Color.FromArgb(244, 244, 244);

        public SqadXlsxViewSheetBuilder(string sheetName, int headerRowsCount)
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

        private void FormatRows(ExcelWorksheet sheet)
        {
            var firstDataRowIndex = _headerRowsCount + 1;
            var allDataCells = sheet.Cells[firstDataRowIndex, 1, sheet.Dimension.Rows, sheet.Dimension.Columns];
            allDataCells.Style.Fill.PatternType = ExcelFillStyle.Solid;
            allDataCells.Style.Fill.BackgroundColor.SetColor(Color.White);
            allDataCells.Style.Border.BorderAround(ExcelBorderStyle.None);

            var tableLegendCells = sheet.Cells[firstDataRowIndex, 1, sheet.Dimension.Rows, FirstColumnsToMerge];
            tableLegendCells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

            var percentsCells = sheet.Cells[firstDataRowIndex, 1, sheet.Dimension.Rows, 1];
            percentsCells.Style.Fill.PatternType = ExcelFillStyle.Solid;
            percentsCells.Style.Fill.BackgroundColor.SetColor(_dataTotalsBackgroundColor);
            percentsCells.Style.Numberformat.Format = "0 %";

            var dataCells = sheet.Cells[firstDataRowIndex, FirstColumnsToMerge + 1, sheet.Dimension.Rows,
                                        sheet.Dimension.Columns];
            dataCells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            dataCells.Style.Numberformat.Format = "#,###";

            FormatTotalDataColumns(sheet, firstDataRowIndex);
        }

        private void FormatTotalDataColumns(ExcelWorksheet sheet, int firstDataRowIndex)
        {
            for (var cellIndex = FirstColumnsToMerge + 1; cellIndex <= sheet.Dimension.Columns; cellIndex++)
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

            var mergedCells = worksheet.Cells[1, 1, _headerRowsCount, FirstColumnsToMerge];
            mergedCells.Merge = true;
            SetBordersToCells(mergedCells);

            for (var rowIndex = 1; rowIndex <= _headerRowsCount - 1; rowIndex++)
            {
                var startColumnIndex = FirstColumnsToMerge + 1;
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

            for (var columnIndex = FirstColumnsToMerge + 1; columnIndex <= worksheet.Dimension.Columns; columnIndex++)
            {
                SetBordersToCells(worksheet.Cells[_headerRowsCount, columnIndex]);
            }
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