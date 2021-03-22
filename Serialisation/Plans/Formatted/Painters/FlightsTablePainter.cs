using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Drawing;
using System.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using SQAD.MTNext.Business.Models.Core.Currency;
using SQAD.MTNext.Business.Models.FlowChart.DataModels;
using WebApiContrib.Formatting.Xlsx.Serialisation.Plans.Formatted.Helpers;
using WebApiContrib.Formatting.Xlsx.Serialisation.Plans.Formatted.Models;

namespace WebApiContrib.Formatting.Xlsx.Serialisation.Plans.Formatted.Painters
{
    internal class FlightsTablePainter
    {
        private const int HEADER_ROW_INDEX = 2;
        private const string DEFAULT_COLUMN_NAME = "No Formula";
        private readonly Dictionary<int, CurrencyModel> _currencies;
        private readonly Dictionary<int, RowDefinition> _planRows;

        private readonly ExcelWorksheet _worksheet;

        public FlightsTablePainter(ExcelWorksheet worksheet,
                                   Dictionary<int, CurrencyModel> currencies,
                                   Dictionary<int, RowDefinition> planRows)
        {
            _worksheet = worksheet;
            _currencies = currencies;
            _planRows = planRows;
        }

        public int DrawFlightsTable(ChartData chartData)
        {
            try
            {
                return DrawFlightsTableUnsafe(chartData);
            }
            catch (Exception ex)
            {
                return 1;
            }
        }

        private int DrawFlightsTableUnsafe(ChartData chartData)
        {
            var pCount = _planRows.Count;
            if (chartData.Objects.Cell != null)
            {
                chartData.Objects.Cell.RemoveAll(c => CellAddress.GetRowIndexByCoords(c.Coordinates) > pCount);
            }
            var maxColumnIndex = 0;
            var maxRowIndex = 0;

            var columnsLookup = chartData.LeftTableColumns.ToDictionary(x => int.Parse(x.Key) + 2);
            columnsLookup.Add(1,
                              new LeftTableColumnValue
                              {
                                  Label = "#",
                                  Appearances = new Appearance
                                                {
                                                    TextAlign = ""
                                                }
                              });

            var separateCells = new Dictionary<string, CellAddress>();
            foreach (var cell in chartData.Objects.Cell ?? new List<ObjectCell>())
            {
                if (cell?.Coordinates == null)
                {
                    continue;
                }

                var address = new CellAddress(cell.Coordinates);
                if (!address.IsFlightsTableAddress)
                {
                    continue;
                }

                address.Cell = cell;
                var key = $"{address.RowIndex}-{address.ColumnIndex}";
                if (!separateCells.ContainsKey(key))
                {
                    separateCells.Add(key, address);
                }
            }

            foreach (var tableCell in chartData.Cells ?? new List<TextValue>())
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

                var rowDefinition = _planRows.GetValueOrDefault(cellAddress.RowIndex);

                var startRowIndex = rowDefinition.StartExcelRowIndex;
                var endRowIndex = rowDefinition.EndExcelRowIndex;
                var cell = _worksheet.Cells[startRowIndex,
                                            cellAddress.ColumnIndex,
                                            endRowIndex,
                                            cellAddress.ColumnIndex];

                var column = columnsLookup.GetValueOrDefault(cellAddress.ColumnIndex);
                var objectCell = separateCells.GetValueOrDefault($"{cellAddress.RowIndex}-{cellAddress.ColumnIndex}")
                                              ?.Cell;

                var cellAppearance = GetAppearance(column, objectCell);

                cell.Merge = true;
                cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                var firstCell = _worksheet.Cells[cell.Start.Address];

                cellAppearance.FillValue(tableCell.FormattedValue != null ? tableCell.FormattedValue : tableCell.Value, firstCell, _currencies, false);
            }

            for (var headerColumnIndex = 1; headerColumnIndex <= maxColumnIndex; headerColumnIndex++)
            {
                var labelCell = _worksheet.Cells[HEADER_ROW_INDEX, headerColumnIndex];
                labelCell.Value = columnsLookup.TryGetValue(headerColumnIndex, out var column)
                                      ? column.Label
                                      : DEFAULT_COLUMN_NAME;
            }

            FormatFlightsTable(maxColumnIndex, columnsLookup);

            foreach (var cellAddress in separateCells.Values)
            {
                var rowDefinition = _planRows.GetValueOrDefault(cellAddress.RowIndex);

                var column = columnsLookup.GetValueOrDefault(cellAddress.ColumnIndex);
                var cellAppearance = GetAppearance(column, cellAddress.Cell);
                var range = _worksheet.Cells[rowDefinition.StartExcelRowIndex,
                                             cellAddress.ColumnIndex,
                                             rowDefinition.EndExcelRowIndex,
                                             cellAddress.ColumnIndex];
                FormatRange(range, cellAppearance);
            }

            return maxColumnIndex;
        }

        public void FillRowNumbers(int maxColumnIndex)
        {
            try
            {
                FillRowNumbersUnsafe(maxColumnIndex);
            }
            catch
            {
                // ignored
            }
        }

        private void FillRowNumbersUnsafe(int maxColumnIndex)
        {
            var estimatedNumberColumnWidth = 5.0;
            var maxRow = _planRows.Keys.Max();

            for (var columnIndex = 1; columnIndex <= maxColumnIndex; columnIndex++)
            {
                var rowNumber = 1;
                for (var rowIndex = 1; rowIndex <= maxRow; rowIndex++)
                {
                    var rowDefinition = _planRows.GetValueOrDefault(rowIndex);

                    var cells = _worksheet.Cells[rowDefinition.StartExcelRowIndex,
                                                 columnIndex,
                                                 rowDefinition.EndExcelRowIndex,
                                                 columnIndex];
                    if (cells.Count() > 1)
                    {
                        cells.Merge = true;
                    }

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

        private void FormatFlightsTable(int maxColumnIndex, Dictionary<int, LeftTableColumnValue> tableColumns)
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
                var leftTableColumn = tableColumns.GetValueOrDefault(columnIndex);

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

                if (leftTableColumn == null)
                {
                    labelCell.Style.Font.Color.SetColor(Colors.DayHeaderHolidayFontColor);
                }
                else
                {
                    var appearance = GetAppearance(leftTableColumn);
                    var columnValues = _worksheet.Cells[HEADER_ROW_INDEX + 2,
                                                        columnIndex,
                                                        _worksheet.Dimension.Rows,
                                                        columnIndex];

                    FormatRange(columnValues, appearance);
                }

                if (columnIndex == maxColumnIndex)
                {
                    column.Style.Border.Right.Style = ExcelBorderStyle.Medium;
                    column.Style.Border.Right.Color.SetColor(Colors.WeekHeaderBorderColor);
                }
            }
        }

        private void FormatRange(ExcelRange range, CellsAppearance appearance)
        {
            AppearanceHelper.SetFromFont(range.Style.Font.SetFromFont, appearance.FontFamily, appearance.FontSize);
            range.Style.Font.Size = appearance.FontSize;
            range.Style.Font.Bold = appearance.Bold;
            range.Style.Font.Italic = appearance.Italic;
            range.Style.Font.Color.SetColor(appearance.TextColor);

            if (appearance.Underline)
            {
                range.Style.Font.UnderLine = true;
                range.Style.Font.UnderLineType = ExcelUnderLineType.Single;
            }
            else
            {
                range.Style.Font.UnderLine = false;
            }

            if (appearance.UseBackColor)
            {
                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(appearance.BackgroundColor);
            }
            else
            {
                range.Style.Fill.PatternType = ExcelFillStyle.None;
            }

            switch (appearance.TextAlignment)
            {
                case eTextAlignment.Left:
                    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    break;
                case eTextAlignment.Right:
                    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    break;
                default:
                    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    break;
            }

            switch (appearance.TextVerticalAlignment)
            {
                case eTextAnchoringType.Top:
                    range.Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    break;
                case eTextAnchoringType.Bottom:
                    range.Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;
                    break;
                default:
                    range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    break;
            }
        }

        private CellsAppearance GetAppearance(LeftTableColumnValue column, ObjectCell cell = null)
        {
            Appearance appearance = null;
            if (column != null)
            {
                appearance = column.Appearances ?? new Appearance();
                appearance.TextAlign = appearance.TextAlign ?? "flex-start";

                if (column.Options?.GetValueOrDefault("currency") is long currency)
                {
                    appearance.UseCurrencySymbol = true;
                    appearance.CurrencySymbol = (int) currency;
                    appearance.FloatingPointAccuracy = 2;
                }
            }

            var cellAppearance = AppearanceHelper.GetAppearance(cell?.Appearance, appearance);

            return cellAppearance;
        }
    }

    public class CellAddress
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

            RowIndex = int.Parse(rawAddress[1]) + 1;
            ColumnIndex = int.Parse(rawAddress[2]) + 2;
        }

        public CellAddress(Coordinates coords)
        {
            IsFlightsTableAddress = coords.Type == FLIGHTS_ADDRESS_MARKER;
            if (!IsFlightsTableAddress)
            {
                return;
            }

            RowIndex = GetRowIndexByCoords(coords);//coords.StartY + 1;
            ColumnIndex = GetColIndexByCoords(coords);//coords.X + 2;
        }

        public static int GetRowIndexByCoords(Coordinates coords)
        {
            return coords.StartY + 1;
        }

        public static int GetColIndexByCoords(Coordinates coords)
        {
            return coords.X + 2;
        }


        public bool IsFlightsTableAddress { get; }

        public int RowIndex { get; }
        public int ColumnIndex { get; }
        public ObjectCell Cell { get; set; }
    }
}