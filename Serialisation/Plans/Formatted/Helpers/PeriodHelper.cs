using SQAD.MTNext.Business.Models.Common.Enums;
using SQAD.MTNext.Business.Models.Core.Calendar;
using SQAD.MTNext.Business.Models.FlowChart.Plan;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using SQAD.MTNext.Business.Models.FlowChart.Export;

namespace WebApiContrib.Formatting.Xlsx.Serialisation.Plans.Formatted.Helpers
{
    //NOTE: not used - client calendar structure is dirty, so use pre-computed periods from PlanCreator
    internal class PeriodHelper
    {
        private const int STANDARD_CALENDAR_ID = -1;
        private const int BROADCAST_CALENDAR_ID = 0;

        private readonly DateTime _planStartDate;
        private readonly DateTime _planEndDate;

        private readonly List<CalendarStructureModel> _calendarStructures;
        private readonly CalendarStructureModel _calendarStructure;
        private readonly int _fiscalPeriodId;

        private readonly CustomPeriodsType? _calendarType;
        private readonly WeekLabel _weekLabel;
        private readonly int _customWeekLabel;

        public PeriodHelper(PlanModel plan,
                            ClientCalendarTypeModel calendarType,
                            List<CalendarStructureModel> calendarStructures)
        {
            _planStartDate = plan.StartDate;
            _planEndDate = _planStartDate.AddDays(plan.Duration);
            _fiscalPeriodId = plan.FiscalPeriodID ?? 0;

            _calendarStructures = calendarStructures;
            _calendarStructure = _calendarStructures.FirstOrDefault(x => x.Id == plan.CalendarTypeID);

            switch (plan.WeekLabel)
            {
                case "CW":
                    _weekLabel = WeekLabel.CalendarWeek;
                    break;
                case "CRWC":
                    _weekLabel = WeekLabel.CustomRunningWeekCount;
                    _customWeekLabel = plan.CustomWeekLabelNumber == 0 ? 1 : plan.CustomWeekLabelNumber;
                    break;
                case "RWC":
                    _weekLabel = WeekLabel.DefaultRunningWeekCount;
                    _customWeekLabel = 1;
                    break;
            }

            switch (plan.CalendarTypeID)
            {
                case BROADCAST_CALENDAR_ID:
                    _calendarType = CustomPeriodsType.Broadcast;
                    break;
                case STANDARD_CALENDAR_ID:
                    _calendarType = CustomPeriodsType.Standard;
                    break;
                default:
                {
                    if (plan.CalendarTypeID.HasValue)
                    {
                        _calendarType = (CustomPeriodsType) calendarType.CustomPeriods;
                    }

                    break;
                }
            }
        }

        public ICollection<CalendarSpan> Build()
        {
            switch (_calendarStructure?.Id)
            {
                case null:
                case STANDARD_CALENDAR_ID:
                case BROADCAST_CALENDAR_ID:
                    return BuildWeeks(_planStartDate, _planEndDate);
                    //var result = new List<CalendarSpan>();
                    //var years = _calendarStructure.Items.Where(x => _planStartDate >= x.StartDate
                    //                                                && _planEndDate <= x.EndDate);
                    //foreach (var year in years)
                    //{
                    //    if (!year.StartDate.HasValue
                    //        || !year.EndDate.HasValue)
                    //    {
                    //        continue;
                    //    }

                    //    result.AddRange(BuildWeeks(year.StartDate.Value, year.EndDate.Value));
                    //}

                    //return result;
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
            var endDate = period.EndDate.Value.AddDays(-1);

            //switch (_calendarType)
            //{
            //    case CustomPeriodsType.Broadcast:
            //        return BuildBroadcastWeeks(startDate, endDate);
            //    case CustomPeriodsType.Standard:
            //        return BuildStandardWeeks(startDate, endDate);
            //}

            return BuildWeeks(startDate, endDate);
        }

        private List<CalendarSpan> BuildBroadcastWeeks(DateTime startDate, DateTime endDate)
        {
            var result = new List<CalendarSpan>();

            var currentWeekLabel = _customWeekLabel;
            if (_weekLabel == WeekLabel.CalendarWeek)
            {
                var startOfYear = new DateTime(startDate.Year, 1, 1);
                while (startOfYear.DayOfWeek != DayOfWeek.Monday)
                {
                    startOfYear = startOfYear.AddDays(-1);
                }

                var startOfCurrentWeek = startDate;
                while (startOfYear.DayOfWeek != DayOfWeek.Monday)
                {
                    startOfCurrentWeek = startOfCurrentWeek.AddDays(-1);
                }

                currentWeekLabel = (int) (startOfCurrentWeek - startOfYear).TotalDays / 7 + 1;
            }

            var previousStartDate = startDate;
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
                               Week = currentWeekLabel,
                               StartDate = previousStartDate,
                               EndDate = currentDate,
                               Name = $"{currentWeekLabel} W"
                           };
                week.Spans = BuildDays(week);

                result.Add(week);

                currentWeekLabel++;
                previousStartDate = currentDate.AddDays(1);
            }

            return result;
        }

        private List<CalendarSpan> BuildWeeks(DateTime startDate, DateTime endDate)
        {
            var result = new List<CalendarSpan>();

            var currentWeekLabel = _customWeekLabel;
            if (_weekLabel == WeekLabel.CalendarWeek)
            {
                if (_calendarType == CustomPeriodsType.Broadcast)
                {
                    var defaultBroadcast = _calendarStructures.First(x => x.Id == BROADCAST_CALENDAR_ID);
                    var currentYear = defaultBroadcast.Items
                                                      .FirstOrDefault(x => startDate >= x.StartDate
                                                                           && startDate <= x.EndDate);

                    var startOfYear = currentYear?.StartDate ?? startDate;

                    var startOfCurrentWeek = startDate;
                    while (startOfCurrentWeek.DayOfWeek != DayOfWeek.Monday)
                    {
                        startOfCurrentWeek = startOfCurrentWeek.AddDays(-1);
                    }

                    currentWeekLabel = (int)(startOfCurrentWeek - startOfYear).TotalDays / 7 + 1;
                }
                else
                {
                    var calendar = new GregorianCalendar(GregorianCalendarTypes.USEnglish);
                    currentWeekLabel = calendar.GetWeekOfYear(startDate, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
                }
                //var calendar = new GregorianCalendar(GregorianCalendarTypes.USEnglish);
                //currentWeekLabel = calendar.GetWeekOfYear(startDate, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
            }

            if (_weekLabel == WeekLabel.CalendarWeek)
            {
                var previousStartDate = startDate;
                for (var currentDate = startDate; currentDate <= endDate; currentDate = currentDate.AddDays(1))
                {
                    //var endOfWeek = _calendarType == CustomPeriodsType.Broadcast
                    //                    ? DayOfWeek.Sunday
                    //                    : DayOfWeek.Saturday;
                    
                    if (currentDate.DayOfWeek != DayOfWeek.Sunday && currentDate != endDate)
                    {
                        continue;
                    }
                    
                    var week = new CalendarSpan
                               {
                                   Year = currentDate.Year,
                                   Month = currentDate.Month,
                                   Week = currentWeekLabel,
                                   StartDate = previousStartDate,
                                   EndDate = currentDate,
                                   Name = $"{currentWeekLabel} W"
                               };
                    week.Spans = BuildDays(week);

                    result.Add(week);

                    currentWeekLabel++;
                    previousStartDate = currentDate.AddDays(1);
                }
            }
            else
            {
                for (var currentDate = startDate; currentDate <= endDate; currentDate = currentDate.AddDays(7))
                {
                    var currentEndDate = currentDate.AddDays(6);
                    currentEndDate = currentEndDate > endDate ? endDate : currentEndDate;

                    var week = new CalendarSpan
                               {
                                   Year = currentDate.Year,
                                   Month = currentDate.Month,
                                   Week = currentWeekLabel,
                                   StartDate = currentDate,
                                   EndDate = currentEndDate,
                                   Name = $"{currentWeekLabel} W"
                               };
                    week.Spans = BuildDays(week);
                    result.Add(week);

                    currentWeekLabel++;
                }
            }

            return result;
        }

        private List<CalendarSpan> BuildStandardWeeks(DateTime startDate, DateTime endDate)
        {
            var result = new List<CalendarSpan>();

            var currentWeekLabel = _customWeekLabel;
            if (_weekLabel == WeekLabel.CalendarWeek)
            {
                var calendar = new GregorianCalendar(GregorianCalendarTypes.USEnglish);
                currentWeekLabel = calendar.GetWeekOfYear(startDate, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
            }

            var previousStartDate = startDate;
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
                               Week = currentWeekLabel,
                               StartDate = previousStartDate,
                               EndDate = currentDate,
                               Name = $"{currentWeekLabel} W"
                           };
                week.Spans = BuildDays(week);

                result.Add(week);

                currentWeekLabel++;
                previousStartDate = currentDate.AddDays(1);
            }

            return result;
        }

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