using System;

namespace ExcelAddIn1.PricerObjects
{
    public static class UniversalDateTime
    {
        public static string ConvertFromTimestampToString(double timestamp)
        {
            var origin = new DateTime(1970, 1, 1, 0, 0, 0, 0);
            return string.Format("{0:yyyyMMdd}", origin.AddSeconds(timestamp));
        }


        public static DateTime ConvertFromTimestamp(double timestamp)
        {
            var origin = new DateTime(1970, 1, 1, 0, 0, 0, 0);
            return origin.AddSeconds(timestamp);
        }

        public static double ConvertToTimestamp(this DateTime date)
        {
            var origin = new DateTime(1970, 1, 1, 0, 0, 0, 0);
            var diff = date - origin;
            return Math.Floor(diff.TotalSeconds);
        }
    }
}