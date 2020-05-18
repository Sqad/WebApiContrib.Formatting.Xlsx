using SQAD.MTNext.Business.Models.Core.Calendar;
using SQAD.MTNext.Business.Models.FlowChart.Plan;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using SQAD.MTNext.Business.Models.Common.Enums;

namespace WebApiContrib.Formatting.Xlsx.Serialisation.Plans.Formatted.Helpers
{
    internal class PeriodHelper
    {
        private const int STANDARD_CALENDAR_ID = -1;
        private const int BROADCAST_CALENDAR_ID = 0;

        private readonly DateTime _planStartDate;
        private readonly DateTime _planEndDate;
        private readonly bool _useIdAsYear;

        private readonly CalendarStructureModel _calendarStructure;
        private readonly int _fiscalPeriodId;

        private readonly CustomPeriodsType? _calendarType;
        private readonly WeekLabel _weekLabel;

        public PeriodHelper(PlanModel plan,
                            ClientCalendarTypeModel calendarType,
                            CalendarStructureModel calendarStructure)
        {
            _planStartDate = plan.StartDate;
            _planEndDate = _planStartDate.AddDays(plan.Duration);
            _fiscalPeriodId = plan.FiscalPeriodID ?? 0;

            _calendarStructure = calendarStructure;

            switch (plan.WeekLabel)
            {
                case "CW":
                    _weekLabel = WeekLabel.CalendarWeek;
                    break;
                case "CRWC":
                    _weekLabel = WeekLabel.CustomRunningWeekCount;
                    break;
                case "RWC":
                    _weekLabel = WeekLabel.DefaultRunningWeekCount;
                    break;
            }

            switch (plan.CalendarTypeID)
            {
                case BROADCAST_CALENDAR_ID:
                    _calendarType = CustomPeriodsType.Broadcast;
                    _useIdAsYear = true;
                    break;
                case STANDARD_CALENDAR_ID:
                    _calendarType = CustomPeriodsType.Standard;
                    _useIdAsYear = true;
                    break;
                default:
                {
                    if (plan.CalendarTypeID.HasValue)
                    {
                        _calendarType = (CustomPeriodsType) calendarType.CustomPeriods;
                    }
                    else
                    {
                        _useIdAsYear = true;
                    }

                    break;
                }
            }
        }

        public ICollection<CalendarSpan> Build()
        {
            //if (!_calendarType.HasValue)
            //{
            //    var result = new List<CalendarSpan>();

            //    var currentDate = _planStartDate;
            //    var currentMonth = new CalendarSpan
            //                       {
            //                           Year = currentDate.Year,
            //                           Month = currentDate.Month,
            //                           StartDate = currentDate,
            //                           Spans = new List<CalendarSpan>(),
            //                           Name = $"{new DateTime(currentDate.Year, currentDate.Month, 1):MMMM} {currentDate.Year}"
            //                       };
            //    do
            //    {
            //        var month = currentDate.Month;
            //        if (currentMonth.Month != month)
            //        {
            //            currentMonth.EndDate = currentDate.AddDays(-1);
            //            result.Add(currentMonth);

            //            currentMonth = new CalendarSpan
            //                           {
            //                               Year = currentDate.Year,
            //                               Month = currentDate.Month,
            //                               StartDate = currentDate,
            //                               Spans = new List<CalendarSpan>(),
            //                               Name = $"{new DateTime(currentDate.Year, currentDate.Month, 1):MMMM} {currentDate.Year}"
            //                           };
            //        }

            //        currentDate = currentDate.AddDays(1);
            //    } while (currentDate <= _planEndDate);

            //    currentMonth.EndDate = currentDate.AddDays(-1);
            //    result.Add(currentMonth);

            //    var weekNumber = 1;
            //    foreach (var month in result)
            //    {
            //        month.Spans = BuildStandardWeeks(month, ref weekNumber);
            //    }

            //    return result;
            //}

            switch (_calendarStructure?.Id)
            {
                case null:
                    return BuildStandardWeeks(_planStartDate, _planEndDate);
                case STANDARD_CALENDAR_ID:
                case BROADCAST_CALENDAR_ID:
                    var result = new List<CalendarSpan>();
                    var years = _calendarStructure.Items.Where(x => _planStartDate >= x.StartDate
                                                                    && _planEndDate <= x.EndDate);
                    foreach (var year in years)
                    {
                        if (!year.StartDate.HasValue
                            || !year.EndDate.HasValue)
                        {
                            continue;
                        }

                        result.AddRange(_calendarStructure.Id == STANDARD_CALENDAR_ID
                                            ? BuildStandardWeeks(year.StartDate.Value, year.EndDate.Value)
                                            : BuildBroadcastWeeks(year.StartDate.Value, year.EndDate.Value));
                    }

                    return result;
            }

            return BuildCalendar();
        }

        private ICollection<CalendarSpan> BuildCalendar()
        {
            var period = _calendarStructure.Items.FirstOrDefault(x => x.Id == _fiscalPeriodId);
            if (period?.StartDate == null || !period.EndDate.HasValue)
            {
                return new List<CalendarSpan>();
            }

            var startDate = period.StartDate.Value;
            var endDate = period.EndDate.Value;

            switch (_calendarType)
            {
                case CustomPeriodsType.Broadcast:
                    return BuildBroadcastWeeks(startDate, endDate);
                case CustomPeriodsType.Standard:
                    return BuildStandardWeeks(startDate, endDate);
            }

            return BuildStandardWeeks(startDate, endDate);
        }

        //private ICollection<CalendarSpan> BuildMonths()
        //{
        //    var result = new List<CalendarSpan>();

        //    var years = _calendarStructure.Items.Where(x => _planStartDate >= x.StartDate
        //                                                    && _planEndDate <= x.EndDate);

        //    var weekNumber = 1;
        //    foreach (var year in years)
        //    {
        //        var yearNumber = _useIdAsYear ? year.Id : year.FiscalYear;

        //        foreach (var quarter in year.Items)
        //        {
        //            foreach (var month in quarter.Items)
        //            {
        //                if (!month.StartDate.HasValue
        //                    || !month.EndDate.HasValue)
        //                {
        //                    continue;
        //                }

        //                var monthSpan = new CalendarSpan
        //                                {
        //                                    Year = yearNumber,
        //                                    Month = month.Id,
        //                                    StartDate = month.StartDate.Value,
        //                                    EndDate = month.EndDate.Value,
        //                                    Name = $"{new DateTime(yearNumber, month.Id, 1):MMMM} {yearNumber}"
        //                                };
        //                monthSpan.Spans = BuildWeeks(monthSpan, ref weekNumber);

        //                result.Add(monthSpan);
        //            }
        //        }
        //    }

        //    return result;
        //}

        //private List<CalendarSpan> BuildWeeks(CalendarSpan month, ref int weekNumber)
        //{
        //    var result = new List<CalendarSpan>();

        //    switch (_calendarType)
        //    {
        //        case CustomPeriodsType.Broadcast:
        //            return BuildBroadcastWeeks(month, ref weekNumber);
        //        case CustomPeriodsType.Standard:
        //            return BuildStandardWeeks(month, ref weekNumber);
        //        //case CustomPeriodsType.Custom://todo
        //        //    break;
        //        //case CustomPeriodsType.Freeform://todo
        //        //    break;
        //        default:
        //            throw new InvalidEnumArgumentException(
        //                $"{nameof(CustomPeriodsType)} has invalid value: {_calendarType}");
        //    }

        //    return result;
        //}

        private List<CalendarSpan> BuildBroadcastWeeks(DateTime startDate, DateTime endDate)
        {
            var result = new List<CalendarSpan>();

            var previousStartDate = startDate;
            var weekNumber = 1;
            for (var currentDate = startDate; currentDate <= endDate; currentDate = currentDate.AddDays(1))
            {
                if (currentDate.DayOfWeek != DayOfWeek.Sunday)
                {
                    continue;
                }

                var week = new CalendarSpan
                           {
                               Year = currentDate.Year,
                               Month = currentDate.Month,
                               Week = weekNumber,
                               StartDate = previousStartDate,
                               EndDate = currentDate,
                               Name = $"{weekNumber} W"
                           };
                week.Spans = BuildDays(week);

                result.Add(week);

                weekNumber++;
                previousStartDate = currentDate.AddDays(1);
            }

            return result;
        }

        private List<CalendarSpan> BuildStandardWeeks(DateTime startDate, DateTime endDate)
        {
            var result = new List<CalendarSpan>();

            var previousStartDate = startDate;
            var weekNumber = 1;
            for (var currentDate = startDate; currentDate <= endDate; currentDate = currentDate.AddDays(1))
            {
                if (currentDate.DayOfWeek != DayOfWeek.Saturday)
                {
                    continue;
                }

                var week = new CalendarSpan
                           {
                               Year = currentDate.Year,
                               Month = currentDate.Month,
                               Week = weekNumber,
                               StartDate = previousStartDate,
                               EndDate = currentDate,
                               Name = $"{weekNumber} W"
                           };
                week.Spans = BuildDays(week);

                result.Add(week);

                weekNumber++;
                previousStartDate = currentDate.AddDays(1);
            }

            return result;
        }

        //private List<CalendarSpan> BuildBroadcastWeeks(CalendarSpan month, ref int weekNumber)
        //{
        //    var result = new List<CalendarSpan>();

        //    var currentStartDate = month.StartDate;
        //    var currentEndDate = currentStartDate.AddDays(6);
        //    while (currentEndDate <= month.EndDate)
        //    {
        //        var week = new CalendarSpan
        //                   {
        //                       Year = month.Year,
        //                       Month = month.Month,
        //                       Week = weekNumber,
        //                       StartDate = currentStartDate,
        //                       EndDate = currentEndDate,
        //                       Name = $"{weekNumber} W"
        //                   };

        //        week.Spans = BuildDays(week);
        //        result.Add(week);

        //        weekNumber++;
        //        currentStartDate = currentEndDate.AddDays(1);
        //        currentEndDate = currentStartDate.AddDays(6);
        //    }

        //    return result;
        //}

        //private List<CalendarSpan> BuildStandardWeeks(CalendarSpan month, ref int weekNumber)
        //{
        //    var result = new List<CalendarSpan>();

        //    var currentDate = month.StartDate;

        //    var startWeekDate = currentDate;
        //    while (currentDate <= month.EndDate)
        //    {
        //        if (currentDate.DayOfWeek == DayOfWeek.Sunday || currentDate == month.EndDate)
        //        {
        //            var week = new CalendarSpan
        //                       {
        //                           Year = month.Year,
        //                           Month = month.Month,
        //                           Week = weekNumber, //todo
        //                           StartDate = startWeekDate,
        //                           EndDate = currentDate,
        //                           Name = $"{weekNumber} W"
        //                       };
        //            week.Spans = BuildDays(week);
        //            result.Add(week);

        //            weekNumber++;
        //            startWeekDate = currentDate.AddDays(1);
        //        }

        //        currentDate = currentDate.AddDays(1);
        //    }

        //    return result;
        //}

        private List<CalendarSpan> BuildDays(CalendarSpan week)
        {
            var result = new List<CalendarSpan>();

            var currentDate = week.StartDate;
            do
            {
                result.Add(new CalendarSpan
                           {
                               Year = week.Year,
                               Month = week.Month,
                               Week = week.Week,
                               StartDate = currentDate,
                               EndDate = currentDate,
                               Day = currentDate.Day,
                               Name = currentDate.Day.ToString()
                           });

                currentDate = currentDate.AddDays(1);
            } while (currentDate <= week.EndDate);

            return result;
        }
    }

    internal class CalendarSpan
    {
        public CalendarSpan()
        {
            Spans = new List<CalendarSpan>();
        }

        public string Name { get; set; }
        public int Year { get; set; }
        public int Month { get; set; }
        public int Week { get; set; }
        public int Day { get; set; }

        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }

        public List<CalendarSpan> Spans { get; set; }
    }

    internal enum WeekLabel
    {
        CalendarWeek,
        CustomRunningWeekCount,
        DefaultRunningWeekCount
    }
}