using System.Drawing;

namespace WebApiContrib.Formatting.Xlsx.Serialisation.Plans.Formatted.Helpers
{
    internal static class Colors
    {
        public static readonly Color WeekHeaderBorderColor = ColorTranslator.FromHtml("#e7e7e7");
        public static readonly Color WeekHeaderFontColor = ColorTranslator.FromHtml("#a8a8a8");
        public static readonly Color DayHeaderBackgroundColor = ColorTranslator.FromHtml("#f6f6f6");
        public static readonly Color DayHeaderHolidayFontColor = ColorTranslator.FromHtml("#999999");
        public static readonly Color DayHeaderBorderColor = ColorTranslator.FromHtml("#fefefe");
        public static readonly Color HolidayColumnBackgroundColor = ColorTranslator.FromHtml("#f7f7f7");

        public static readonly Color DefaultFlightBackgroundColor = ColorTranslator.FromHtml("#00acdb");
    }
}
