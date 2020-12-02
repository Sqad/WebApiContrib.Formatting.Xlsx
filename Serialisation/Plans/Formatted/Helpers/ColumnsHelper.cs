using System;
using System.Collections.Generic;
using System.Linq;
using SQAD.MTNext.Business.Models.FlowChart.DataModels;
using SQAD.MTNext.Business.Models.FlowChart.Enums;

namespace WebApiContrib.Formatting.Xlsx.Serialisation.Plans.Formatted.Helpers
{
    internal class ColumnsHelper
    {
        private readonly FormattedPlanViewMode _viewMode;
        private readonly List<HeaderColumn> _columns;

        public ColumnsHelper(FormattedPlanViewMode viewMode, HeaderColumns columns)
        {
            _viewMode = viewMode;

            switch (_viewMode)
            {
                case FormattedPlanViewMode.Daily:
                    _columns = columns.Daily;
                    break;
                case FormattedPlanViewMode.Weekly:
                    _columns = columns.Weekly;
                    break;
                case FormattedPlanViewMode.Monthly:
                    throw new NotSupportedException("Monthly view is not supported");
            }
        }

        public ICollection<CalendarSpan> Build()
        {
            switch (_viewMode)
            {
                case FormattedPlanViewMode.Daily:
                    return BuildDaily();
                case FormattedPlanViewMode.Weekly:
                    return BuildWeekly();
                default:
                    throw new NotSupportedException("Monthly view is not supported");
            }
        }

        private ICollection<CalendarSpan> BuildWeekly()
        {
            var result = new List<CalendarSpan>();

            foreach (var column in _columns.OrderBy(x => x.Index))
            {
                var week = new CalendarSpan
                           {
                               Year = column.Year,
                               Month = column.Month + 1,
                               Week = column.Week,
                               StartDate = column.StartDate,
                               EndDate = column.EndDate.AddDays(-1),
                               Name = $"{column.Week} W"
                           };
                week.Spans = BuildDays(week);

                result.Add(week);
            }

            return result;
        }

        private ICollection<CalendarSpan> BuildDaily()
        {
            var result = new List<CalendarSpan>();
            foreach (var weekGroup in _columns.GroupBy(x => new { x.Year, x.Week} ).OrderBy(x => x.Key.Year).ThenBy(x => x.Key.Week))
            {
                var weekNumber = weekGroup.Key.Week;
                var start = weekGroup.First();
                var end = weekGroup.Last();

                var week = new CalendarSpan
                           {
                               Year = start.Year,
                               Month = start.Month + 1,
                               Week = weekNumber,
                               StartDate = start.StartDate,
                               EndDate = end.StartDate,
                               Name = $"{weekNumber} W"
                           };
                week.Spans = BuildDays(week);

                result.Add(week);
            }

            //CalendarSpan currentWeek = null;
            //var previousWeek = 1;

            //foreach (var column in _columns.OrderBy(x => x.Index))
            //{
            //    if (currentWeek == null)
            //    {
            //        currentWeek = new CalendarSpan
            //        {
            //            Year = column.Year,
            //            Month = column.Month,
            //            Week = column.Week,
            //            StartDate = column.StartDate,
            //            Name = $"{column.Week} W"
            //        };
            //        previousWeek = column.Week;
            //    }

            //    if (column.Week == previousWeek)
            //    {
            //        continue;
            //    }

            //    currentWeek.EndDate = column.StartDate;
            //    currentWeek.Spans = BuildDays(currentWeek);

            //    result.Add(currentWeek);
            //}

            return result;
        }

        private static List<CalendarSpan> BuildDays(CalendarSpan week)
        {
            var result = new List<CalendarSpan>();

            var currentDate = week.StartDate;
            do
            {
                result.Add(new CalendarSpan
                           {
                               Year = week.Year,
                               Month = week.Month + 1,
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
}