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
        private const string DEFAULT_COLUMN_NAME = "No Formula";

        private readonly ExcelWorksheet _worksheet;

        public FlightsTablePainter(ExcelWorksheet worksheet)
        {
            _worksheet = worksheet;
        }

        public int DrawFlightsTable(ChartData chartData)
        {
            var maxColumnIndex = 0;
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

                var cell = _worksheet.Cells[cellAddress.RowIndex, cellAddress.ColumnIndex];
                cell.Value = tableCell.Value;

                if (tableCell.Value is DateTime)
                {
                    cell.Style.Numberformat.Format = "mm-dd-yy";
                }
            }

            var columnsLookup = chartData.LeftTableColumns.ToDictionary(x => int.Parse(x.Key) + 1);
            for (var headerColumnIndex = 1; headerColumnIndex <= maxColumnIndex; headerColumnIndex++)
            {
                var labelCell = _worksheet.Cells[HEADER_ROW_INDEX, headerColumnIndex];
                labelCell.Value = columnsLookup.TryGetValue(headerColumnIndex, out var column)
                                      ? column.Label
                                      : DEFAULT_COLUMN_NAME;

                //var addressCell = _worksheet.Cells[HEADER_ROW_INDEX + 1, headerColumnIndex];
                //addressCell.Value = GetExcelColumnName(headerColumnIndex);
            }

            FormatFlightsTable(maxColumnIndex, columnsLookup.Keys.ToImmutableHashSet());

            return maxColumnIndex;
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

                //var addressCell = _worksheet.Cells[HEADER_ROW_INDEX + 1, columnIndex];
                //addressCell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                //addressCell.Style.Font.Color.SetColor(Colors.DayHeaderHolidayFontColor);

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

                RowIndex = int.Parse(rawAddress[1]) + 2 + HEADER_ROW_INDEX;
                ColumnIndex = int.Parse(rawAddress[2]) + 1;
            }

            public bool IsFlightsTableAddress { get; }

            public int RowIndex { get; }
            public int ColumnIndex { get; }
        }
    }
}