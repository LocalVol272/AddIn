using System;
using System.Collections.Generic;

namespace ExcelAddIn1.PricerObjects
{
    public static class UniversalDateTime
    {
        public static string ConvertFromTimestampToString(this double timestamp)
        {
            var origin = new DateTime(1970, 1, 1, 0, 0, 0, 0);
            return string.Format("{0:yyyyMMdd}", origin.AddSeconds(timestamp));
        }

        public static List<string> ConvertFromTimestampToString(this List<double> timestamp)
        {
            var res = new List<string>();

            timestamp.ForEach(x => res.Add(x.ConvertFromTimestampToString()));
            return res;
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