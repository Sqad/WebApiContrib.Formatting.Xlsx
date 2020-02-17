using OfficeOpenXml;
using OfficeOpenXml.Style;
using SQAD.MTNext.Business.Models.FlowChart.DataModels;
using System;
using System.Collections.Generic;
using System.Drawing;
using WebApiContrib.Formatting.Xlsx.Serialisation.Plans.Formatted.Helpers;

namespace WebApiContrib.Formatting.Xlsx.Serialisation.Plans.Formatted.Painters
{
    internal class FlightPainter
    {
        private readonly ExcelWorksheet _worksheet;
        private readonly bool _daily;
        private readonly int _rowsOffset;
        private readonly Dictionary<DateTime, int> _columnsLookup;

        public FlightPainter(ExcelWorksheet worksheet,
                             bool daily,
                             int rowsOffset,
                             Dictionary<DateTime, int> columnsLookup)
        {
            _worksheet = worksheet;
            _daily = daily;
            _rowsOffset = rowsOffset;
            _columnsLookup = columnsLookup;
        }

        public void DrawFlight(Flight flight, VehicleModel vehicle)
        {
            var rowIndex = (flight.RowIndex ?? 1) + _rowsOffset;

            var startDate = flight.StartDate.Date;
            var endDate = flight.EndDate.AddDays(-1).Date;

            var startColumn = _columnsLookup[startDate];
            var endColumn = _columnsLookup[endDate];

            var flightCells = _worksheet.Cells[rowIndex, startColumn, rowIndex, endColumn];

            FormatFlight(flightCells, flight, vehicle);

            flightCells.Value = flight.Name;
        }

        private static void FormatFlight(ExcelRange cells, Flight flight, VehicleModel vehicle)
        {
            cells.Merge = true;

            var appearance = GetAppearance(flight, vehicle);

            cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
            cells.Style.Fill.BackgroundColor.SetColor(appearance.BackgroundColor);
        }

        private static CellsAppearance GetAppearance(Flight flight, VehicleModel vehicle)
        {
            var appearance = GetMergedAppearance(flight, vehicle);
            var cellsAppearance = new CellsAppearance();

            if (appearance.UseBackColor ?? false)
            {
                cellsAppearance.BackgroundColor = ColorTranslator.FromHtml(appearance.BackColor);
            }
            else
            {
                cellsAppearance.BackgroundColor = Colors.DefaultFlightBackgroundColor;
            }

            return cellsAppearance;
        }

        private static Appearance GetMergedAppearance(Flight flight, VehicleModel vehicle)
        {
            var appearance = new Appearance();

            if (vehicle != null)
            {
                appearance.UseBackColor = vehicle.Appearance.UseBackColor;
                appearance.BackColor = vehicle.Appearance.BackColor;
            }

            appearance.UseBackColor = flight.Appearance.UseBackColor ?? appearance.UseBackColor;
            appearance.BackColor = flight.Appearance.BackColor ?? appearance.BackColor;

            return appearance;
        }

        private class CellsAppearance
        {
            public Color BackgroundColor { get; set; }
        }
    }
}