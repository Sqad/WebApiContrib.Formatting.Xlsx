using System;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using SQAD.MTNext.Business.Models.FlowChart.DataModels;
using SQAD.MTNext.Business.Models.FlowChart.Enums;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.Drawing;
using SQAD.MTNext.Business.Models.Core.Currency;
using WebApiContrib.Formatting.Xlsx.Serialisation.Plans.Formatted.Helpers;
using WebApiContrib.Formatting.Xlsx.Serialisation.Plans.Formatted.Models;

namespace WebApiContrib.Formatting.Xlsx.Serialisation.Plans.Formatted.Painters
{
    internal class FormulasPainter
    {
        private readonly ExcelWorksheet _worksheet;
        private readonly int _columnsOffset;
        private readonly FormattedPlanViewMode _viewMode;
        private readonly Dictionary<int, CurrencyModel> _currencies;
        private readonly Dictionary<DateTime, int> _columnsLookup;
        private readonly Dictionary<int, RowDefinition> _planRows;

        public FormulasPainter(ExcelWorksheet worksheet,
                               int columnsOffset,
                               FormattedPlanViewMode viewMode,
                               Dictionary<int, CurrencyModel> currencies,
                               Dictionary<DateTime, int> columnsLookup,
                               Dictionary<int, RowDefinition> planRows)
        {
            _worksheet = worksheet;
            _columnsOffset = columnsOffset;
            _viewMode = viewMode;
            _currencies = currencies;
            _columnsLookup = columnsLookup;
            _planRows = planRows;
        }

        public void DrawFormulas(ChartData chartData)
        {
            try
            {
                DrawFormulasUnsafe(chartData);
            }
            catch
            {
                // ignored
            }
        }

        private void DrawFormulasUnsafe(ChartData chartData)
        {
            var subtotalRowsLookup = (chartData.Objects
                                              .SubtotalRows ?? new List<SubtotalRow>())
                                              .GroupBy(x => x.Id)
                                              .ToDictionary(x => x.Key, x => x.First());

            var cellsLookup = BuildCellsLookup(chartData);

            string formulaType = null;
            switch (_viewMode)
            {
                case FormattedPlanViewMode.Daily:
                    formulaType = "formulaDaily";
                    break;
                case FormattedPlanViewMode.Weekly:
                    formulaType = "formulaWeekly";
                    break;
            }

            if (formulaType == null)
            {
                return;
            }


            var formulasLookup = (chartData.Objects
                                    .Formulas ?? new List<Formula>())
                                    .Where(x => x.FormulaType == formulaType)
                                    .GroupBy(x => x.RowIndex)
                                    .ToDictionary(x => x.Key,
                                                  x => x.GroupBy(y => y.ColumnIndex + _columnsOffset)
                                                        .ToDictionary(y => y.Key, y => y.First()));

            foreach (var currentCells in cellsLookup)
            {
                var rowIndex = currentCells.Key;
                var rowDefinition = _planRows.GetValueOrDefault(rowIndex);

                var cells = currentCells.Value;

                var formulas = formulasLookup.GetValueOrDefault(rowIndex);
                foreach (var cell in cells)
                {
                    var startColumnIndex = cell.ColumnIndex;
                    var formula = formulas?.GetValueOrDefault(startColumnIndex);
                    if (formula == null)
                    {
                        continue;
                    }

                    startColumnIndex = _columnsLookup[formula.StartDate.Date];
                    var endColumnIndex = _columnsLookup[formula.EndDate.AddDays(-1).Date];

                    var ranges = new List<ExcelRange>();

                    var previousStartColumn = startColumnIndex;
                    for (var columnIndex = startColumnIndex; columnIndex <= endColumnIndex; columnIndex++)
                    {
                        var excelCell = _worksheet.Cells[rowDefinition.PrimaryExcelRowIndex, columnIndex];
                        if (!excelCell.Merge)
                        {
                            if (endColumnIndex == columnIndex)
                            {
                                ranges.Add(_worksheet.Cells[rowDefinition.PrimaryExcelRowIndex,
                                                            previousStartColumn,
                                                            rowDefinition.PrimaryExcelRowIndex,
                                                            columnIndex]);
                            }

                            continue;
                        }

                        if (previousStartColumn == columnIndex)
                        {
                            previousStartColumn++;
                            continue;
                        }

                        ranges.Add(_worksheet.Cells[rowDefinition.PrimaryExcelRowIndex,
                                                    previousStartColumn,
                                                    rowDefinition.PrimaryExcelRowIndex,
                                                    columnIndex - 1]);

                        previousStartColumn = columnIndex + 1;
                    }

                    if (!ranges.Any())
                    {
                        continue;
                    }

                    Appearance subtotalRowAppearance = null;
                    if (formula.SubtotalId.HasValue)
                    {
                        var subtotalRow = subtotalRowsLookup.GetValueOrDefault(formula.SubtotalId.Value);
                        subtotalRowAppearance = subtotalRow?.Appearance;
                    }

                    var appearance = AppearanceHelper.GetAppearance(formula.Appearance, subtotalRowAppearance);

                    var firstRange = ranges.First();
                    var firstCellAddress = firstRange.Address;
                    var firstCell = _worksheet.Cells[firstCellAddress];

                    appearance.FillValue(cell.Value.FormattedValue != null? cell.Value.FormattedValue : cell.Value.Value
                           , firstCell, _currencies);

                    foreach (var range in ranges)
                    {
                        ApplyAppearance(range, appearance);
                    }
                }
            }
        }

        private static void ApplyAppearance(ExcelRange range, CellsAppearance appearance)
        {
            range.Merge = true;
            range.Style.ShrinkToFit = true;
            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            var backColor = appearance.UseBackColor
                                ? appearance.BackgroundColor 
                                : Colors.DefaultFormulaBackgroundColor;

            range.Style.Fill.PatternType = ExcelFillStyle.Solid;
            range.Style.Fill.BackgroundColor.SetColor(backColor);

            range.Style.Border.BorderAround(ExcelBorderStyle.Thin,
                                            appearance.UseCellBorderColor
                                                ? appearance.CellBorderColor
                                                : Colors.DefaultFormulaBorderColor);

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

        private Dictionary<int, SubtotalCellAddress[]> BuildCellsLookup(ChartData chartData)
        {
            var subtotalCells = new Dictionary<int, List<SubtotalCellAddress>>();

            foreach (var cell in chartData.Cells ?? new List<TextValue>())
            {
                var cellAddress = new SubtotalCellAddress(cell.Key, _columnsOffset, _viewMode);
                if (!cellAddress.IsSubtotalCellAddress)
                {
                    continue;
                }

                cellAddress.Value = cell;
                if (!subtotalCells.ContainsKey(cellAddress.RowIndex))
                {
                    subtotalCells.Add(cellAddress.RowIndex, new List<SubtotalCellAddress>());
                }

                subtotalCells[cellAddress.RowIndex].Add(cellAddress);
            }

            return subtotalCells.ToDictionary(x => x.Key, x => x.Value.OrderBy(y => y.ColumnIndex).ToArray());
        }

        private class SubtotalCellAddress
        {
            private const string WEEKLY_ADDRESS_MARKER = "W";
            private const string DAILY_ADDRESS_MARKER = "D";
            private const string FLIGHTS_ADDRESS_SEPARATOR = ":";

            public SubtotalCellAddress(string key, int columnsOffset, FormattedPlanViewMode viewMode)
            {
                var rawAddress = key.Split(FLIGHTS_ADDRESS_SEPARATOR);

                switch (viewMode)
                {
                    case FormattedPlanViewMode.Daily:
                        IsSubtotalCellAddress = rawAddress[0] == DAILY_ADDRESS_MARKER;
                        break;
                    case FormattedPlanViewMode.Weekly:
                        IsSubtotalCellAddress = rawAddress[0] == WEEKLY_ADDRESS_MARKER;
                        break;
                }

                if (!IsSubtotalCellAddress)
                {
                    return;
                }

                RowIndex = int.Parse(rawAddress[1]) + 1;
                ColumnIndex = int.Parse(rawAddress[2]) + columnsOffset;
            }

            public bool IsSubtotalCellAddress { get; }
            public int RowIndex { get; }
            public int ColumnIndex { get; }
            public TextValue Value { get; set; }
        }
    }
}