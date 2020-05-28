using Newtonsoft.Json;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using SQAD.MTNext.Business.Models.FlowChart.DataModels;
using SQAD.MTNext.Business.Models.FlowChart.Enums;
using SQAD.MTNext.Business.Models.FlowChart.Plan;
using SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.Base;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using WebApiContrib.Formatting.Xlsx.Serialisation.Plans.Formatted.Helpers;
using WebApiContrib.Formatting.Xlsx.Serialisation.Plans.Formatted.Models;
using WebApiContrib.Formatting.Xlsx.Serialisation.Plans.Formatted.Painters;

namespace WebApiContrib.Formatting.Xlsx.Serialisation.Plans.Formatted
{
    internal class FormattedPlanSheetBuilder : SqadXlsxSheetBuilderBase
    {
        private const int HEADER_HEIGHT = 3;

        private readonly ExportPlanRequest _exportPlanRequest;
        private readonly ColumnsHelper _columnsHelper;
        private readonly FormattedPlanViewMode _viewMode;
        private readonly ChartData _chartData;

        private Dictionary<DateTime, int> _columnsLookup;
        private int _flightsTableWidth;

        private readonly Dictionary<int, RowDefinition> _planRows;

        public FormattedPlanSheetBuilder(string sheetName, ExportPlanRequest exportPlanRequest)
            : base(sheetName)
        {
            _exportPlanRequest = exportPlanRequest;
            _viewMode = exportPlanRequest.ViewMode;

            _chartData = JsonConvert.DeserializeObject<ChartData>(exportPlanRequest.Chart.Version.JsonData,
                                                                  new JsonSerializerSettings
                                                                  {
                                                                      StringEscapeHandling =
                                                                          StringEscapeHandling.EscapeHtml,
                                                                      NullValueHandling = NullValueHandling.Ignore,
                                                                      MissingMemberHandling =
                                                                          MissingMemberHandling.Ignore
                                                                  });

            if (_chartData.Columns == null)
            {
                throw new ApplicationException("Please re-save plan before export");
            }

            _columnsHelper = new ColumnsHelper(_viewMode, _chartData.Columns);

            _planRows = new Dictionary<int, RowDefinition>();
        }

        protected override void CompileSheet(ExcelWorksheet worksheet, DataTable table)
        {
            RowHeightsHelper.FillRowHeights(_planRows, _chartData);

            var flightsTablePainter = new FlightsTablePainter(worksheet, _exportPlanRequest.Currencies, _planRows);
            _flightsTableWidth = flightsTablePainter.DrawFlightsTable(_chartData);

            FillCalendarHeader(worksheet);
            FillGrid(worksheet);
            FillFormulas(worksheet);
            FillCaptions(worksheet);
            FillShapes(worksheet);

            flightsTablePainter.FillRowNumbers(_flightsTableWidth);

            for (var rowIndex = HEADER_HEIGHT + 1; rowIndex <= worksheet.Dimension.Rows; rowIndex++)
            {
                var row = worksheet.Row(rowIndex);
                var firstCell = worksheet.Cells[rowIndex, 1];
                if (firstCell.Merge)
                {
                    continue;
                }

                row.Height = 30;
            }

            FillPictures(worksheet);
        }

        private void FillCalendarHeader(ExcelWorksheet worksheet)
        {
            var calendarSpans = _columnsHelper.Build();

            _columnsLookup = new Dictionary<DateTime, int>();

            const int monthRowIndex = 1;
            const int dayRowIndex = monthRowIndex + 1;
            const int weekRowIndex = monthRowIndex + 2;

            var monthColumnStartIndex = _flightsTableWidth + 1;
            var currentMonthColumnIndex = monthColumnStartIndex;
            DateTime? lastMonth = null;

            var weekColumnStartIndex = monthColumnStartIndex;
            var dayColumnIndex = weekColumnStartIndex;

            foreach (var week in calendarSpans)
            {
                worksheet.Cells[weekRowIndex, weekColumnStartIndex].Value = week.Name;

                var days = _viewMode == FormattedPlanViewMode.Daily ? week.Spans : week.Spans.Take(1);
                foreach (var day in days)
                {
                    _columnsLookup.Add(day.StartDate, dayColumnIndex);

                    var dayCell = worksheet.Cells[dayRowIndex, dayColumnIndex];
                    dayCell.Value = day.Day;

                    var isHoliday = _viewMode == FormattedPlanViewMode.Daily
                                    && (day.StartDate.DayOfWeek == DayOfWeek.Saturday
                                        || day.StartDate.DayOfWeek == DayOfWeek.Sunday);
                    FormatColumn(worksheet.Column(dayColumnIndex), isHoliday);
                    FormatDayCells(dayCell, isHoliday);

                    dayColumnIndex++;
                }

                if (_viewMode != FormattedPlanViewMode.Daily)
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

                var currentMonth = new DateTime(week.Year, week.Month, 1);
                if (lastMonth == null)
                {
                    lastMonth = currentMonth;
                    continue;
                }

                if (currentMonth == lastMonth)
                {
                    currentMonthColumnIndex = dayColumnIndex;
                    continue;
                }

                var monthCells = worksheet.Cells[monthRowIndex,
                                                 monthColumnStartIndex,
                                                 monthRowIndex,
                                                 currentMonthColumnIndex
                                                 - ((monthColumnStartIndex >= currentMonthColumnIndex) ? 0 : 1)];
                monthCells.Value = $"{lastMonth.Value:MMMM} {lastMonth.Value.Year}";
                
                FormatMonthCells(monthCells);

                lastMonth = currentMonth;
                monthColumnStartIndex = currentMonthColumnIndex;
            }

            if (lastMonth != null)
            {
                var monthCells = worksheet.Cells[monthRowIndex,
                                                 monthColumnStartIndex,
                                                 monthRowIndex,
                                                 currentMonthColumnIndex 
                                                   - ((monthColumnStartIndex >= currentMonthColumnIndex)? 0 : 1)];
                monthCells.Value = $"{lastMonth.Value:MMMM} {lastMonth.Value.Year}";

                FormatMonthCells(monthCells);
            }

            worksheet.View.FreezePanes(HEADER_HEIGHT + 1, 1);
        }

        private void FillFormulas(ExcelWorksheet worksheet)
        {
            var formulasPainter = new FormulasPainter(worksheet,
                                                      _flightsTableWidth,
                                                      _viewMode,
                                                      _exportPlanRequest.Currencies,
                                                      _columnsLookup,
                                                      _planRows);
            formulasPainter.DrawFormulas(_chartData);
        }

        private void FillGrid(ExcelWorksheet worksheet)
        {
            int? mRowIndex = 0;

            var flightPainter = new FlightPainter(worksheet, _columnsLookup, _planRows, _exportPlanRequest.Currencies);
            
            if (_chartData.Objects.Flights != null)
            {
                List<Flight> flights = _chartData.Objects.Flights;
                List<FlightHelper> intersectFlights = new List<FlightHelper>();
                if (_viewMode == FormattedPlanViewMode.Weekly)
                {
                    if (flights.Any())
                    {
                      flights = flights.OrderBy(x => x.RowIndex).ThenBy(x => x.StartDate).ToList();
                      mRowIndex = flights.First().RowIndex;
                    }
                }
            
                FlightHelper lastFlightHelper = null;

                foreach (var flight in flights)
                {
                    VehicleModel vehicle = null;
                    if (flight.VehicleID.HasValue)
                    {
                        vehicle = _chartData.Vehicles.FirstOrDefault(x => x.ID == flight.VehicleID);
                    }

                    int flightRowIndex = -1;

                    if (_viewMode == FormattedPlanViewMode.Weekly)
                    {
                        if (lastFlightHelper != null)
                        {
                            if (flight != lastFlightHelper.Flight)
                            {
                                continue;
                            }
                        }

                        int index = flights.IndexOf(flight);
                        Flight nextFlight = flights.ElementAtOrDefault(index + 1);
                        int currIndex = 0;
                        int nextIndex = 0;
                        while ((nextFlight != null) && (nextFlight.RowIndex == mRowIndex) && (currIndex == nextIndex))
                        {
                            currIndex = _columnsLookup[flight.EndDate.AddDays(-1).Date];
                            nextIndex = _columnsLookup[nextFlight.StartDate];
                            if (currIndex == nextIndex)
                            {
                                if (!intersectFlights.Any())
                                {
                                    var firstFlightHelper = new FlightHelper(flight);
                                    if (lastFlightHelper != null)
                                    {
                                        firstFlightHelper.StartCorrection = lastFlightHelper.StartCorrection;
                                        firstFlightHelper.EndCorrection = lastFlightHelper.EndCorrection;
                                    }
                                        intersectFlights.Add(firstFlightHelper);
                                    
                                }
                                intersectFlights.Add(new FlightHelper(nextFlight));
                            }
                            index += 1;
                            nextFlight = flights.ElementAtOrDefault(index + 1);
                        }

                        if ((nextFlight != null) && (nextFlight.RowIndex > mRowIndex))
                        {
                            mRowIndex = nextFlight.RowIndex;
                        }

                        if (intersectFlights.Any())
                        {
                            int daysInWeek = 0;
                            int flightIndex = 0;
                            int maxDaysFlightIndex = 0;
                            
                            foreach (var fl in intersectFlights)
                            {
                                int currDaysInWeek = 1;
                                int weekNumber = 0;
                                
                                if (flightIndex == 0)
                                {
                                    weekNumber = _columnsLookup[fl.Flight.EndDate.AddDays(-1).Date];
                                    DateTime dateTime = fl.Flight.EndDate.AddDays(-1 * currDaysInWeek - 1).Date;
                                    while ((_columnsLookup[dateTime] == weekNumber) && (dateTime >= fl.Flight.StartDate))
                                    {
                                        currDaysInWeek += 1;
                                        dateTime = fl.Flight.EndDate.AddDays(-1 * currDaysInWeek - 1).Date;
                                    }
                                    currDaysInWeek -= 1;
                                }
                                else
                                {
                                    weekNumber = _columnsLookup[fl.Flight.StartDate.Date];
                                    DateTime dateTime = fl.Flight.StartDate.AddDays(currDaysInWeek).Date;
                                    while ((_columnsLookup[dateTime] == weekNumber) && (dateTime <= fl.Flight.EndDate.AddDays(-1).Date))
                                    {
                                        currDaysInWeek += 1;
                                        dateTime = fl.Flight.StartDate.AddDays(currDaysInWeek).Date;
                                    }
                                }
                                if (daysInWeek < currDaysInWeek)
                                {
                                    daysInWeek = currDaysInWeek;
                                    maxDaysFlightIndex = flightIndex;
                                }

                                flightIndex += 1;
                                
                            }
                            foreach (var fl in intersectFlights)
                            {
                                if (intersectFlights.IndexOf(fl) < maxDaysFlightIndex)
                                {
                                    fl.EndCorrection -= 1;
                                }

                                if (intersectFlights.IndexOf(fl) > maxDaysFlightIndex)
                                {
                                    fl.StartCorrection += 1;
                                }

                                if (fl != intersectFlights.Last())
                                {
                                    flightRowIndex = flightPainter.DrawFlight(fl, vehicle);
                                }
                            }
                            lastFlightHelper = intersectFlights.Last();
                            currIndex = 0;
                            nextIndex = 0;
                            intersectFlights.Clear();
                        }
                        else 
                        {
                            if (lastFlightHelper != null)
                            {
                                flightRowIndex = flightPainter.DrawFlight(lastFlightHelper, vehicle);
                                lastFlightHelper = null;
                            }
                            else
                            {
                               flightRowIndex = flightPainter.DrawFlight(new FlightHelper(flight), vehicle);
                            }
                        }
                    }
                    else
                    {
                        flightRowIndex = flightPainter.DrawFlight(new FlightHelper(flight), vehicle);
                    }
                }
            }
        }

        private void FillCaptions(ExcelWorksheet worksheet)
        {
            var captionPainter = new CaptionPainter(worksheet, _columnsLookup, _planRows);
            foreach (var caption in _chartData.Objects.Texts ?? new List<Text>())
            {
                captionPainter.DrawCaption(caption);
            }
        }

        private void FillShapes(ExcelWorksheet worksheet)
        {
            var shapePainter = new ShapePainter(worksheet, _columnsLookup, _planRows);
            foreach (var shape in _chartData.Objects.Shapes ?? new List<Shape>())
            {
                shapePainter.DrawShape(shape);
            }
        }

        private void FillPictures(ExcelWorksheet worksheet)
        {
            var picturesPainter = new PicturesPainter(worksheet, _columnsLookup, _planRows);
            picturesPainter.DrawPictures(_chartData.Objects.Pictures ?? new List<Picture>(), _viewMode);
        }

        private static void FormatMonthCells(ExcelRangeBase cells)
        {
            cells.Merge = true;
            cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            cells.Style.Font.Size = 14;

            cells.Style.Fill.PatternType = ExcelFillStyle.None;

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

            cells.Style.Fill.PatternType = ExcelFillStyle.None;

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
            column.Width = 5.71;
        }
    }
}