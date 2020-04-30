using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Drawing;
using OfficeOpenXml;
using SQAD.MTNext.Business.Models.FlowChart.DataModels;
using System.Linq;
using System.Text;
using OfficeOpenXml.Style;
using WebApiContrib.Formatting.Xlsx.Serialisation.Plans.Formatted.Helpers;

namespace WebApiContrib.Formatting.Xlsx.Serialisation.Plans.Formatted.Painters
{
    internal class FlightsTablePainter
    {
        private const int HEADER_ROW_INDEX = 2;
        private const int ROW_MULTIPLIER = 3;
        private const string DEFAULT_COLUMN_NAME = "No Formula";

        private readonly ExcelWorksheet _worksheet;

        public FlightsTablePainter(ExcelWorksheet worksheet)
        {
            _worksheet = worksheet;
        }

        public (int maxColumnIndex, int maxRowIndex) DrawFlightsTable(ChartData chartData)
        {
            var maxColumnIndex = 0;
            var maxRowIndex = 0;
            foreach (var tableCell in chartData.Cells)
            {
                var cellAddress = new CellAddress(tableCell.Key);
                if (!cellAddress.IsFlightsTableAddress)
                {
                    continue;
                }

                if (cellAddress.ColumnIndex > maxColumnIndex)
                {
                    maxColumnIndex = cellAddress.ColumnIndex;
                }

                if (cellAddress.RowIndex > maxRowIndex)
                {
                    maxRowIndex = cellAddress.RowIndex;
                }

                var startRowIndex = cellAddress.RowIndex - 1;
                var endRowIndex = cellAddress.RowIndex + 1;
                var cell = _worksheet.Cells[startRowIndex, cellAddress.ColumnIndex, endRowIndex,
                                            cellAddress.ColumnIndex];

                cell.Merge = true;
                cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                cell.Value = tableCell.Value;

                if (tableCell.Value is DateTime)
                {
                    cell.Style.Numberformat.Format = "mm-dd-yy";
                }
            }

            var columnsLookup = chartData.LeftTableColumns.ToDictionary(x => int.Parse(x.Key) + 2);
            columnsLookup.Add(1, new TextValue
                                 {
                                     Label = "#"
                                 });

            for (var headerColumnIndex = 1; headerColumnIndex <= maxColumnIndex; headerColumnIndex++)
            {
                var labelCell = _worksheet.Cells[HEADER_ROW_INDEX, headerColumnIndex];
                labelCell.Value = columnsLookup.TryGetValue(headerColumnIndex, out var column)
                                      ? column.Label
                                      : DEFAULT_COLUMN_NAME;
            }

            FormatFlightsTable(maxColumnIndex, columnsLookup.Keys.ToImmutableHashSet());

            return (maxColumnIndex, maxRowIndex);
        }

        public void FillRowNumbers(int maxRowIndex, int maxColumnIndex)
        {
            var estimatedNumberColumnWidth = 5.0;

            for (var columnIndex = 1; columnIndex <= maxColumnIndex; columnIndex++)
            {
                var rowNumber = 1;
                for (var rowIndex = 5; rowIndex <= maxRowIndex; rowIndex += ROW_MULTIPLIER)
                {
                    var cells = _worksheet.Cells[rowIndex - 1, columnIndex, rowIndex + 1, columnIndex];
                    cells.Merge = true;

                    if (columnIndex != 1)
                    {
                        continue;
                    }

                    cells.Value = rowNumber;
                    cells.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    var columnWidth = Math.Floor(Math.Log10(rowNumber) + 1);
                    if (estimatedNumberColumnWidth < columnWidth)
                    {
                        estimatedNumberColumnWidth = columnWidth;
                    }

                    rowNumber++;
                }
            }

            var column = _worksheet.Column(1);
            column.Width = estimatedNumberColumnWidth;
        }

        private void FormatFlightsTable(int maxColumnIndex, ICollection<int> columnsWithLabel)
        {
            var emptyCells = _worksheet.Cells[1, 1, 1, maxColumnIndex];
            emptyCells.Style.Fill.PatternType = ExcelFillStyle.Solid;
            emptyCells.Style.Fill.BackgroundColor.SetColor(Color.White);

            for (var columnIndex = 1; columnIndex <= maxColumnIndex; columnIndex++)
            {
                var column = _worksheet.Column(columnIndex);
                column.AutoFit(12);
                column.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                var labelCell = _worksheet.Cells[HEADER_ROW_INDEX, columnIndex];
                var isLabelDefined = columnsWithLabel.Contains(columnIndex);

                labelCell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                labelCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                labelCell.Style.Fill.BackgroundColor.SetColor(Colors.DayHeaderBackgroundColor);

                labelCell.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                labelCell.Style.Border.Left.Color.SetColor(Colors.DayHeaderBorderColor);

                if (columnIndex != maxColumnIndex)
                {
                    labelCell.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    labelCell.Style.Border.Right.Color.SetColor(Colors.DayHeaderBorderColor);
                }

                if (!isLabelDefined)
                {
                    labelCell.Style.Font.Color.SetColor(Colors.DayHeaderHolidayFontColor);
                }

                if (columnIndex == maxColumnIndex)
                {
                    column.Style.Border.Right.Style = ExcelBorderStyle.Medium;
                    column.Style.Border.Right.Color.SetColor(Colors.WeekHeaderBorderColor);
                }
            }
        }

        private static string GetExcelColumnName(int columnIndex)
        {
            var dividend = columnIndex;
            var columnNameBuilder = new StringBuilder();

            while (dividend > 0)
            {
                var modulo = (dividend - 1) % 26;
                columnNameBuilder.Insert(0, Convert.ToChar(65 + modulo));
                dividend = (dividend - modulo) / 26;
            }

            return columnNameBuilder.ToString();
        }

        private class CellAddress
        {
            private const string FLIGHTS_ADDRESS_MARKER = "T";
            private const string FLIGHTS_ADDRESS_SEPARATOR = ":";

            public CellAddress(string key)
            {
                var rawAddress = key.Split(FLIGHTS_ADDRESS_SEPARATOR);

                IsFlightsTableAddress = rawAddress[0] == FLIGHTS_ADDRESS_MARKER;

                if (!IsFlightsTableAddress)
                {
                    return;
                }

                RowIndex = (int.Parse(rawAddress[1]) + 1) * ROW_MULTIPLIER + HEADER_ROW_INDEX;
                ColumnIndex = int.Parse(rawAddress[2]) + 2;
            }

            public bool IsFlightsTableAddress { get; }

            public int RowIndex { get; }
            public int ColumnIndex { get; }
        }
    }
}