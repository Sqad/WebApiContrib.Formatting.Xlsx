using OfficeOpenXml;
using OfficeOpenXml.Style;
using SQAD.MTNext.Business.Models.FlowChart.DataModels;
using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.Drawing;
using WebApiContrib.Formatting.Xlsx.Serialisation.Plans.Formatted.Formulas;
using WebApiContrib.Formatting.Xlsx.Serialisation.Plans.Formatted.Helpers;

namespace WebApiContrib.Formatting.Xlsx.Serialisation.Plans.Formatted.Painters
{
    internal class FlightPainter
    {
        private readonly ExcelWorksheet _worksheet;
        private readonly int _rowsOffset;
        private readonly Dictionary<DateTime, int> _columnsLookup;
        private readonly FormulaParser _formulaParser;

        public FlightPainter(ExcelWorksheet worksheet,
                             int rowsOffset,
                             Dictionary<DateTime, int> columnsLookup)
        {
            _worksheet = worksheet;
            _rowsOffset = rowsOffset;
            _columnsLookup = columnsLookup;

            _formulaParser = new FormulaParser();
        }

        public int DrawFlight(Flight flight,
                              VehicleModel vehicle)
        {
            var rowIndex = ((flight.RowIndex ?? 1) * 3 - 1) + _rowsOffset;

            var startDate = flight.StartDate.Date;
            var endDate = flight.EndDate.AddDays(-1).Date;

            var startColumn = _columnsLookup[startDate];
            var endColumn = _columnsLookup[endDate];

            var appearance = AppearanceHelper.GetAppearance(flight, vehicle);

            var aboveRowIndex = rowIndex - 1;
            DrawAboveCaption(flight, aboveRowIndex, startColumn, endColumn, appearance);

            var flightCells = _worksheet.Cells[rowIndex, startColumn, rowIndex, endColumn];

            var belowRowIndex = rowIndex + 1;
            DrawBelowCaption(flight, belowRowIndex, startColumn, endColumn, appearance);

            FormatFlight(flightCells, appearance);

            flightCells.Value = _formulaParser.GetInsideCaption(flight);

            return rowIndex;
        }

        private static void FormatFlight(ExcelRange cells, CellsAppearance appearance)
        {
            ApplyAppearance(cells, appearance);

            cells.Style.Border.BorderAround(ExcelBorderStyle.Thin, appearance.CellBorderColor);
        }

        private void DrawAboveCaption(Flight flight,
                                      int rowNumber,
                                      int startColumn,
                                      int endColumn,
                                      CellsAppearance appearance)
        {
            var cells = DrawCaption(flight.FlightCaption.Above, flight, rowNumber, startColumn, endColumn, appearance);

            cells.Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;
            ApplyAppearance(cells, appearance);

            cells.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            cells.Style.Border.Top.Color.SetColor(appearance.CellBorderColor);
        }

        private void DrawBelowCaption(Flight flight,
                                      int rowNumber,
                                      int startColumn,
                                      int endColumn,
                                      CellsAppearance appearance)
        {
            var cells = DrawCaption(flight.FlightCaption.Below, flight, rowNumber, startColumn, endColumn, appearance);

            cells.Style.VerticalAlignment = ExcelVerticalAlignment.Top;
            ApplyAppearance(cells, appearance);

            cells.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            cells.Style.Border.Bottom.Color.SetColor(appearance.CellBorderColor);
        }

        private ExcelRange DrawCaption(IReadOnlyCollection<FlightCaptionPosition> captions,
                                       Flight flight,
                                       int rowNumber,
                                       int startColumn,
                                       int endColumn,
                                       CellsAppearance appearance)
        {
            var cells = _worksheet.Cells[rowNumber, startColumn, rowNumber, endColumn];
            cells.Style.WrapText = true;

            cells.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            cells.Style.Border.Right.Color.SetColor(appearance.CellBorderColor);

            cells.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            cells.Style.Border.Left.Color.SetColor(appearance.CellBorderColor);

            if (captions == null || !captions.Any())
            {
                return cells;
            }

            var values = captions.Select(x => _formulaParser.GetValueFromFormula(x.Text, flight)).ToList();
            var captionsText = string.Join('\n', values);

            cells.Value = captionsText;

            var row = _worksheet.Row(rowNumber);
            var neededHeight = 15.0 * values.Count;
            if (row.Height < neededHeight)
            {
                row.Height = neededHeight;
            }

            return cells;
        }

        private static void ApplyAppearance(ExcelRange cells, CellsAppearance appearance)
        {
            cells.Merge = true;

            switch (appearance.TextAlignment)
            {
                case eTextAlignment.Left:
                    cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    break;
                case eTextAlignment.Right:
                    cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    break;
                default:
                    cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    break;
            }

            switch (appearance.TextVerticalAlignment)
            {
                case eTextAnchoringType.Top:
                    cells.Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    break;
                case eTextAnchoringType.Bottom:
                    cells.Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;
                    break;
                default:
                    cells.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    break;
            }

            cells.Style.Font.Color.SetColor(appearance.TextColor);
            cells.Style.Font.Size = appearance.FontSize;
            cells.Style.Font.Bold = appearance.Bold;
            cells.Style.Font.Italic = appearance.Italic;

            if (appearance.Underline)
            {
                cells.Style.Font.UnderLine = appearance.Underline;
                cells.Style.Font.UnderLineType = ExcelUnderLineType.Single;
            }

            cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
            cells.Style.Fill.BackgroundColor.SetColor(appearance.BackgroundColor);
        }
    }
}