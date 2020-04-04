using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace ExcelAddIn1.PricerObjects
{
    internal class Date : IYahooDateFormat
    {
        public Date(int year, int month, [Optional] int day)
        {
            Year = year;
            Month = month;
            Day = day;
        }


        public Date(string date)
        {
            if (date.Length == 8)
            {
                Year = int.Parse(date.Substring(0, 4));
                Month = int.Parse(date.Substring(4, 2));
                Day = int.Parse(date.Substring(6, 2));
            }
            else if (date.Length == 6)
            {
                Year = int.Parse(date.Substring(0, 4));
                Month = int.Parse(date.Substring(4, 2));
            }
            else
            {
                throw new Exception(DataLoaderError.DateFormatError);
            }
        }

        public int Day { get; set; }

        public int Month { get; set; }

        public int Year { get; set; }


        public double ToTimeStamp()
        {
            var dt = new DateTime(Year, Month, Day);
            return dt.ConvertToTimestamp();
        }
    }

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