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

        private readonly CalendarStructureModel _calendarStructure;

        private readonly CustomPeriodsType _calendarType;

        public PeriodHelper(PlanModel plan,
                            ClientCalendarTypeModel calendarType,
                            CalendarStructureModel calendarStructure)
        {
            _planStartDate = plan.StartDate;
            _planEndDate = _planStartDate.AddDays(plan.Duration);

            _calendarStructure = calendarStructure;

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
            if (_calendarStructure == null)
            {
                //todo
                return null;
            }

            var months = BuildMonths();
            return months;
        }

        private ICollection<CalendarSpan> BuildMonths()
        {
            var result = new List<CalendarSpan>();

            var years = _calendarStructure.Items.Where(x => _planStartDate >= x.StartDate
                                                            && _planEndDate <= x.EndDate);

            var weekNumber = 1;
            foreach (var year in years)
            {
                foreach (var quarter in year.Items)
                {
                    foreach (var month in quarter.Items)
                    {
                        if (!month.StartDate.HasValue
                            || !month.EndDate.HasValue)
                        {
                            continue;
                        }

                        var monthSpan = new CalendarSpan
                                        {
                                            Year = year.Id,
                                            Month = month.Id,
                                            StartDate = month.StartDate.Value,
                                            EndDate = month.EndDate.Value,
                                            Name = $"{new DateTime(year.Id, month.Id, 1):MMMM} {year.Id}"
                                        };
                        monthSpan.Spans = BuildWeeks(monthSpan, ref weekNumber);

                        result.Add(monthSpan);
                    }
                }
            }

            return result;
        }

        private List<CalendarSpan> BuildWeeks(CalendarSpan month, ref int weekNumber)
        {
            var result = new List<CalendarSpan>();

            switch (_calendarType)
            {
                case CustomPeriodsType.Broadcast:
                {
                    var currentStartDate = month.StartDate;
                    var currentEndDate = currentStartDate.AddDays(6);
                    while (currentEndDate <= month.EndDate)
                    {
                        var week = new CalendarSpan
                                   {
                                       Year = month.Year,
                                       Month = month.Month,
                                       Week = weekNumber, //todo
                                       StartDate = currentStartDate,
                                       EndDate = currentEndDate,
                                       Name = $"{weekNumber} W" //todo
                                   };

                        week.Spans = BuildDays(week);
                        result.Add(week);

                        weekNumber++;
                        currentStartDate = currentEndDate.AddDays(1);
                        currentEndDate = currentStartDate.AddDays(6);
                    }

                    break;
                }
                case CustomPeriodsType.Standard:
                {
                    var currentDate = month.StartDate;

                    var startWeekDate = currentDate;
                    while (currentDate <= month.EndDate)
                    {
                        if (currentDate.DayOfWeek == DayOfWeek.Sunday || currentDate == month.EndDate)
                        {
                            var week = new CalendarSpan
                                       {
                                           Year = month.Year,
                                           Month = month.Month,
                                           Week = weekNumber, //todo
                                           StartDate = startWeekDate,
                                           EndDate = currentDate,
                                           Name = $"{weekNumber} W"
                                       };
                            week.Spans = BuildDays(week);
                            result.Add(week);

                            weekNumber++;
                            startWeekDate = currentDate.AddDays(1);
                        }

                        currentDate = currentDate.AddDays(1);
                    }

                    break;
                }
                case CustomPeriodsType.Custom://todo
                    break;
                case CustomPeriodsType.Freeform://todo
                    break;
                default:
                    throw new InvalidEnumArgumentException($"{nameof(CustomPeriodsType)} has invalid value: {_calendarType}");
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
}