using System;
using System.Runtime.InteropServices;

namespace ExcelAddIn1.PricerObjects
{
    internal class Date : IYahooDateFormat
    {
        private int day, month, year;

        public int Day
        {
            get => day;
            set => day = value;
        }

        public int Month
        {
            get => month;
            set => month = value;
        }

        public int Year
        {
            get => year;
            set => year = value;
        }

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


        public double ToTimeStamp()
        {
            var dt = new DateTime(Year, Month, Day);
            return dt.ConvertToTimestamp();
        }
    }
}