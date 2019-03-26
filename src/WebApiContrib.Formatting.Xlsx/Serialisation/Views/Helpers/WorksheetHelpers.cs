using System.Collections.Generic;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;

namespace SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Views.Helpers
{
    internal static class WorksheetHelpers
    {
        public const int RowNameColumnIndex = 2;
        public const int RowMeasureColumnIndex = RowNameColumnIndex + 1;
        public const string TotalRowIndicator = "Total";
        
        public static readonly Color TotalsBackgroundColor = Color.FromArgb(227, 236, 248);
        public static readonly Color DataTotalsBackgroundColor = Color.FromArgb(244, 244, 244);
        public static readonly Color HeaderTotalsBackgroundColor = Color.FromArgb(198, 217, 241);
        public static readonly Color HeaderFontColor = Color.FromArgb(47, 79, 79);
        public static readonly Color HeaderBackgroundColor = Color.FromArgb(212, 227, 244);

        public static void FormatDataRows(ExcelWorksheet sheet, int firstDataRowIndex, ICollection<int> totalColumnIndexes)
        {
            var isPreviousRowTotal = false;
            for (var rowIndex = firstDataRowIndex; rowIndex <= sheet.Dimension.Rows; rowIndex++)
            {
                var row = sheet.Cells[rowIndex, 1, rowIndex, sheet.Dimension.Columns];

                var nameCell = sheet.Cells[rowIndex, RowNameColumnIndex];
                if (IsTotalRow(sheet, rowIndex))
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

                if (IsGroupRow(sheet, rowIndex, totalColumnIndexes))
                {
                    FormatGroupRow(row);
                }
            }
        }

        public static void FormatRows(ExcelWorksheet sheet, int firstDataRowIndex, int leftPaneWidth)
        {
            var allDataCells = sheet.Cells[firstDataRowIndex, 1, sheet.Dimension.Rows, sheet.Dimension.Columns];
            allDataCells.Style.Fill.PatternType = ExcelFillStyle.Solid;
            allDataCells.Style.Fill.BackgroundColor.SetColor(Color.White);
            allDataCells.Style.Border.BorderAround(ExcelBorderStyle.None);

            var leftPaneCells = sheet.Cells[firstDataRowIndex, 1, sheet.Dimension.Rows, leftPaneWidth];
            leftPaneCells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

            var percentsCells = sheet.Cells[firstDataRowIndex, 1, sheet.Dimension.Rows, 1];
            percentsCells.Style.Fill.PatternType = ExcelFillStyle.Solid;
            percentsCells.Style.Fill.BackgroundColor.SetColor(DataTotalsBackgroundColor);
            percentsCells.Style.Numberformat.Format = "0 %";

            var beginningCells = sheet.Cells[firstDataRowIndex, leftPaneWidth, sheet.Dimension.Rows, leftPaneWidth];
            FormatNumbers(beginningCells);

            if (sheet.Dimension.Columns <= leftPaneWidth)
            {
                return;
            }

            var dataCells = sheet.Cells[firstDataRowIndex, leftPaneWidth + 1, sheet.Dimension.Rows,
                                        sheet.Dimension.Columns];
            FormatNumbers(dataCells);
        }

        public static void FormatHeader(ExcelWorksheet sheet, int headerRowsCount, ICollection<int> totalColumnIndexes)
        {
            var allHeaderCells = sheet.Cells[1, 1, headerRowsCount, sheet.Dimension.Columns];
            allHeaderCells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            allHeaderCells.Style.VerticalAlignment = ExcelVerticalAlignment.Distributed;

            allHeaderCells.Style.Fill.PatternType = ExcelFillStyle.Solid;
            allHeaderCells.Style.Fill.BackgroundColor.SetColor(HeaderBackgroundColor);

            allHeaderCells.Style.Font.Color.SetColor(HeaderFontColor);

            var firstDataRowIndex = headerRowsCount + 1;
            foreach (var totalColumnIndex in totalColumnIndexes)
            {
                var verticalCellsToMerge = sheet.Cells[1, totalColumnIndex, headerRowsCount, totalColumnIndex];
                verticalCellsToMerge.Value = sheet.Cells[1, totalColumnIndex].Value;
                verticalCellsToMerge.Merge = true;
                verticalCellsToMerge.Style.Fill.PatternType = ExcelFillStyle.Solid;
                verticalCellsToMerge.Style.Fill.BackgroundColor.SetColor(HeaderTotalsBackgroundColor);
                SetBordersToCells(verticalCellsToMerge);

                var totalCells = sheet.Cells[firstDataRowIndex, totalColumnIndex, sheet.Dimension.Rows,
                                             totalColumnIndex];
                totalCells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                totalCells.Style.Fill.BackgroundColor.SetColor(DataTotalsBackgroundColor);
            }
        }

        public static bool IsGroupRow(ExcelWorksheet sheet, int rowIndex, ICollection<int> totalColumnIndexes)
        {
            var isGroupingRow = true;
            for (var x = RowNameColumnIndex + 1; x <= sheet.Dimension.Columns; x++)
            {
                if (sheet.Cells[rowIndex, x].Value == null || totalColumnIndexes.Contains(x))
                {
                    continue;
                }

                isGroupingRow = false;
                break;
            }

            return isGroupingRow;
        }

        private static void FormatTotalRow(ExcelRange row)
        {
            row.Style.Fill.PatternType = ExcelFillStyle.Solid;
            row.Style.Fill.BackgroundColor.SetColor(TotalsBackgroundColor);
        }

        public static bool IsEmptyCell(ExcelRange cell)
        {
            return cell.Value == null
                   || cell.Value as string == string.Empty
                   || cell.Value as string == "-";
        }

        public static bool IsTotalRow(ExcelWorksheet sheet, int rowIndex)
        {
            var nameCell = sheet.Cells[rowIndex, RowNameColumnIndex];
            var percentsCell = sheet.Cells[rowIndex, 1];
            if (nameCell.Value != null
                && nameCell.Value is string name
                && name.Contains(TotalRowIndicator)
                && percentsCell.Value == null)
            {
                return true;
            }

            return false;
        }

        public static bool IsDataRow(ExcelWorksheet sheet, int rowIndex)
        {
            var nameCell = sheet.Cells[rowIndex, RowNameColumnIndex];
            var percentsCell = sheet.Cells[rowIndex, 1];
            if (nameCell.Value != null && percentsCell.Value != null)
            {
                return true;
            }

            return false;
        }

        public static void FormatGroupRow(ExcelRange row)
        {
            row.Style.Font.Bold = true;
        }

        public static void FormatNumbers(ExcelRange cells)
        {
            cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            cells.Style.Numberformat.Format = "#,###";
        }

        public static void SetBordersToCells(ExcelRange cells)
        {
            cells.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.White);
        }
    }
}
