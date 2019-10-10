using System;

namespace SQAD.MTNext.WebApiContrib.Formatting.Xlsx.Serialisation.ApprovalReports.Helpers
{
    public static class Parser
    {
        internal static Func<string, DateTime?> ParseNullableDateTime = val =>
        {
            DateTime value;
            return DateTime.TryParse(val, out value) ? (DateTime?)value : null;
        };

        internal static Func<string, int?> ParseNullableInt = val =>
        {
            int value;
            return int.TryParse(val, out value) ? (int?)value : null;
        };
        internal static Func<string, float?> ParseNullableFloat = val =>
        {
            float value;
            return float.TryParse(val, out value) ? (float?)value : null;
        };
    }
}
