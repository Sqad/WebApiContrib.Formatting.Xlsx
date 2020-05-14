using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using SQAD.MTNext.Business.Models.FlowChart.DataModels;
using System;
using System.Collections.Generic;
using WebApiContrib.Formatting.Xlsx.Serialisation.Plans.Formatted.Formulas;
using WebApiContrib.Formatting.Xlsx.Serialisation.Plans.Formatted.Helpers;
using WebApiContrib.Formatting.Xlsx.Serialisation.Plans.Formatted.Models;

namespace WebApiContrib.Formatting.Xlsx.Serialisation.Plans.Formatted.Painters
{
    internal class FlightPainter
    {
        private readonly ExcelWorksheet _worksheet;
        private readonly Dictionary<DateTime, int> _columnsLookup;
        private readonly Dictionary<int, RowDefinition> _planRows;
        private readonly FormulaParser _formulaParser;

        public FlightPainter(ExcelWorksheet worksheet,
                             Dictionary<DateTime, int> columnsLookup,
                             Dictionary<int, RowDefinition> planRows)
        {
            _worksheet = worksheet;
            _columnsLookup = columnsLookup;
            _planRows = planRows;

            _formulaParser = new FormulaParser();
        }

        public int DrawFlight(FlightHelper flightHelper,
                              VehicleModel vehicle)
        {
            Flight flight = flightHelper.Flight;
            var rowDefinition = _planRows.GetValueOrDefault(flight.RowIndex ?? 0);

            var startDate = flight.StartDate.Date;
            var endDate = flight.EndDate.AddDays(-1).Date;

            var startColumn = _columnsLookup[startDate] + flightHelper.StartCorrection;
            var endColumn = _columnsLookup[endDate] + flightHelper.EndCorrection;

            if (startColumn > endColumn)
            {
                return rowDefinition.EndExcelRowIndex;
            }

            var appearance = AppearanceHelper.GetAppearance(flight, vehicle);

            DrawAboveCaptions(flight, rowDefinition, startColumn, endColumn, appearance);

            var flightCells = _worksheet.Cells[rowDefinition.PrimaryExcelRowIndex,
                                               startColumn,
                                               rowDefinition.PrimaryExcelRowIndex,
                                               endColumn];

            DrawBelowCaptions(flight, rowDefinition, startColumn, endColumn, appearance);

            FormatFlight(flightCells, appearance);

            flightCells.Value = _formulaParser.GetInsideCaption(flight);

            return rowDefinition.EndExcelRowIndex;
        }

        private static void FormatFlight(ExcelRange cells, CellsAppearance appearance)
        {
            ApplyAppearance(cells, appearance);

            cells.Style.Border.BorderAround(ExcelBorderStyle.Thin, appearance.CellBorderColor);
        }

        private void DrawAboveCaptions(Flight flight,
                                       RowDefinition rowDefinition,
                                       int startColumn,
                                       int endColumn,
                                       CellsAppearance appearance)
        {
            var startRowNumber = rowDefinition.StartExcelRowIndex
                                 + (rowDefinition.AboveCount - flight.FlightCaption.Above?.Count ?? 0);
            DrawCaptions(flight.FlightCaption.Above,
                         flight,
                         startRowNumber,
                         startColumn,
                         endColumn,
                         appearance,
                         true);
        }

        private void DrawBelowCaptions(Flight flight,
                                       RowDefinition rowDefinition,
                                       int startColumn,
                                       int endColumn,
                                       CellsAppearance appearance)
        {
            DrawCaptions(flight.FlightCaption.Below,
                         flight,
                         rowDefinition.PrimaryExcelRowIndex + 1,
                         startColumn,
                         endColumn,
                         appearance,
                         false);
        }

        private void DrawCaptions(IReadOnlyCollection<FlightCaptionPosition> captions,
                                  Flight flight,
                                  int startRowNumber,
                                  int startColumn,
                                  int endColumn,
                                  CellsAppearance appearance,
                                  bool drawBorderAbove)
        {
            var current = startRowNumber;
            var index = 0;

            captions = captions ?? new List<FlightCaptionPosition>();

            foreach (var caption in captions)
            {
                var cells = _worksheet.Cells[current, startColumn, current, endColumn];
                cells.Style.WrapText = true;

                cells.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                cells.Style.Border.Right.Color.SetColor(appearance.CellBorderColor);

                cells.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                cells.Style.Border.Left.Color.SetColor(appearance.CellBorderColor);

                if (drawBorderAbove && index == 0)
                {
                    cells.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    cells.Style.Border.Top.Color.SetColor(appearance.CellBorderColor);
                }
                else if (!drawBorderAbove && index == captions.Count - 1)
                {
                    cells.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    cells.Style.Border.Bottom.Color.SetColor(appearance.CellBorderColor);
                }

                var value = caption.Text;//_formulaParser.GetValueFromFormula(caption.Text, flight);

                cells.Value = value;

                //var row = _worksheet.Row(current);
                //var neededHeight = 15.0 * values.Count;
                //if (row.Height < neededHeight)
                //{
                //    row.Height = neededHeight;
                //}
                cells.Merge = true;
                cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                current++;
                index++;
            }
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