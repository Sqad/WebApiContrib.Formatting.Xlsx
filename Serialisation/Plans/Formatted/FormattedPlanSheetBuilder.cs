using System;
using System.Collections.Generic;
using OfficeOpenXml;
using SQAD.MTNext.Business.Models.FlowChart.Plan;
using SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Base;
using System.Data;
using System.Drawing;
using System.Linq;
using Newtonsoft.Json;
using OfficeOpenXml.Style;
using SQAD.MTNext.Business.Models.FlowChart.DataModels;
using WebApiContrib.Formatting.Xlsx.Serialisation.Plans.Formatted.Helpers;
using WebApiContrib.Formatting.Xlsx.Serialisation.Plans.Formatted.Painters;

namespace WebApiContrib.Formatting.Xlsx.Serialisation.Plans.Formatted
{
    internal class FormattedPlanSheetBuilder : SqadXlsxSheetBuilderBase
    {
        private const int HEADER_HEIGHT = 3;

        private readonly ExportPlanRequest _exportPlanRequest;
        private readonly PeriodHelper _periodHelper;
        private readonly bool _daily;
        private readonly ChartData _chartData;

        private Dictionary<DateTime, int> _columnsLookup;
        private int _flightsTableWidth;

        public FormattedPlanSheetBuilder(string sheetName, ExportPlanRequest exportPlanRequest)
            : base(sheetName)
        {
            _exportPlanRequest = exportPlanRequest;
            _periodHelper = new PeriodHelper(exportPlanRequest.Chart.Plan,
                                             exportPlanRequest.ClientCalendarType,
                                             exportPlanRequest.CalendarStructure);

            _daily = exportPlanRequest.Daily;

            _chartData = JsonConvert.DeserializeObject<ChartData>(exportPlanRequest.Chart.Version.JsonData,
                                                                  new JsonSerializerSettings
                                                                  {
                                                                      StringEscapeHandling = StringEscapeHandling.EscapeHtml,
                                                                      NullValueHandling = NullValueHandling.Ignore,
                                                                      MissingMemberHandling = MissingMemberHandling.Ignore
                                                                  });
        }

        protected override void CompileSheet(ExcelWorksheet worksheet, DataTable table)
        {
            var flightsTablePainter = new FlightsTablePainter(worksheet);
            var indexes = flightsTablePainter.DrawFlightsTable(_chartData);
            _flightsTableWidth = indexes.maxColumnIndex;

            FillCalendarHeader(worksheet);
            var maxFlightIndex = FillGrid(worksheet);

            var maxRowIndex = Math.Max(indexes.maxRowIndex, maxFlightIndex);
            flightsTablePainter.FillRowNumbers(maxRowIndex, _flightsTableWidth);
        }

        private void FillCalendarHeader(ExcelWorksheet worksheet)
        {
            var calendarSpans = _periodHelper.Build();

            //var sh = worksheet.Drawings.AddShape("", eShapeStyle.Rect);
            _columnsLookup = new Dictionary<DateTime, int>();

            const int monthRowIndex = 1;
            const int dayRowIndex = monthRowIndex + 1;
            const int weekRowIndex = monthRowIndex + 2;

            var monthColumnStartIndex = _flightsTableWidth + 1;
            var weekColumnStartIndex = monthColumnStartIndex;
            var dayColumnIndex = weekColumnStartIndex;
            foreach (var month in calendarSpans)
            {
                worksheet.Cells[monthRowIndex, monthColumnStartIndex].Value = month.Name;

                foreach (var week in month.Spans)
                {
                    worksheet.Cells[weekRowIndex, weekColumnStartIndex].Value = week.Name;

                    var days = _daily ? week.Spans : week.Spans.Take(1);
                    foreach (var day in days)
                    {
                        _columnsLookup.Add(day.StartDate, dayColumnIndex);

                        var dayCell = worksheet.Cells[dayRowIndex, dayColumnIndex];
                        dayCell.Value = day.Day;

                        var isHoliday = _daily
                                        && (day.StartDate.DayOfWeek == DayOfWeek.Saturday
                                            || day.StartDate.DayOfWeek == DayOfWeek.Sunday);
                        FormatColumn(worksheet.Column(dayColumnIndex), isHoliday);
                        FormatDayCells(dayCell, isHoliday);

                        dayColumnIndex++;
                    }

                    if (!_daily)
                    {
                        foreach (var day in week.Spans.Skip(1))
                        {
                            _columnsLookup.Add(day.StartDate, dayColumnIndex - 1);
                        }
                    }

                    FormatWeekCells(worksheet.Cells[weekRowIndex,
                                                    weekColumnStartIndex,
                                                    weekRowIndex,
                                                    dayColumnIndex - 1]);

                    weekColumnStartIndex = dayColumnIndex;
                }

                FormatMonthCells(worksheet.Cells[monthRowIndex,
                                                 monthColumnStartIndex,
                                                 monthRowIndex,
                                                 dayColumnIndex - 1]);

                monthColumnStartIndex = dayColumnIndex;
            }

            worksheet.View.FreezePanes(HEADER_HEIGHT + 1, 1);
        }

        private int FillGrid(ExcelWorksheet worksheet)
        {
            var maxRowIndex = 0;

            var flightPainter = new FlightPainter(worksheet, _daily, HEADER_HEIGHT, _columnsLookup);
            foreach (var flight in _chartData.Objects.Flights)
            {
                VehicleModel vehicle = null;
                if (flight.VehicleID.HasValue)
                {
                    vehicle = _chartData.Vehicles.FirstOrDefault(x => x.ID == flight.VehicleID);
                }

                var flightRowIndex = flightPainter.DrawFlight(flight, vehicle);
                if (maxRowIndex < flightRowIndex)
                {
                    maxRowIndex = flightRowIndex;
                }
            }

            return maxRowIndex;
        }

        private static void FormatMonthCells(ExcelRangeBase cells)
        {
            cells.Merge = true;
            cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            cells.Style.Font.Size = 14;

            cells.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            cells.Style.Border.Left.Color.SetColor(Colors.WeekHeaderBorderColor);

            cells.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            cells.Style.Border.Right.Color.SetColor(Colors.WeekHeaderBorderColor);
        }

        private static void FormatDayCells(ExcelRangeBase cells, bool holiday)
        {
            cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
            cells.Style.Fill.BackgroundColor.SetColor(Colors.DayHeaderBackgroundColor);

            if (holiday)
            {
                cells.Style.Font.Color.SetColor(Colors.DayHeaderHolidayFontColor);
            }
            
            cells.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            cells.Style.Border.Left.Color.SetColor(Colors.DayHeaderBorderColor);

            cells.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            cells.Style.Border.Right.Color.SetColor(Colors.DayHeaderBorderColor);

            cells.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            cells.Style.Border.Bottom.Color.SetColor(Color.White);
        }

        private static void FormatWeekCells(ExcelRangeBase cells)
        {
            cells.Merge = true;
            cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            cells.Style.Font.Color.SetColor(Colors.WeekHeaderFontColor);

            cells.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            cells.Style.Border.Left.Color.SetColor(Colors.WeekHeaderBorderColor);

            cells.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            cells.Style.Border.Right.Color.SetColor(Colors.WeekHeaderBorderColor);

            cells.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            cells.Style.Border.Bottom.Color.SetColor(Colors.WeekHeaderBorderColor);
        }

        private static void FormatColumn(ExcelColumn column, bool holiday)
        {
            column.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            column.Style.Border.Left.Color.SetColor(Colors.WeekHeaderBorderColor);

            column.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            column.Style.Border.Right.Color.SetColor(Colors.WeekHeaderBorderColor);

            column.Style.Fill.PatternType = ExcelFillStyle.Solid;

            column.Style.Fill.BackgroundColor.SetColor(holiday 
                                                           ? Colors.HolidayColumnBackgroundColor 
                                                           : Color.White);
        }
    }
}