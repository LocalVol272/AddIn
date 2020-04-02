using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

namespace ProjetVolSto.PricerObjects
{
    class Date:IYahooDateFormat
    {
        int day, month, year;
        public int Day { get => day; set => day = value; }
        public int Month { get => month; set => month = value; }
        public int Year { get => year; set => year = value; }
        
        public Date(int year,int month,[Optional] int day)
        {
            Year = year; Month= month; Day = day; 
        }


        public Date(string date)
        {
            if (date.Length == 8)
            {
                Year = Int32.Parse(date.Substring(0, 4));
                Month = Int32.Parse(date.Substring(4, 2));
                Day = Int32.Parse(date.Substring(6, 2));
            }
            else if(date.Length == 6)
            {
                Year = Int32.Parse(date.Substring(0, 4));
                Month = Int32.Parse(date.Substring(4, 2));
            }
            else { throw new Exception(DataLoaderError.DateFormatError);}
        }
        


        public double ToTimeStamp()
        {
            DateTime dt = new DateTime(Year, Month, Day);
            return dt.ConvertToTimestamp();
        }




       

    }

    public static class UniversalDateTime
    {

        public static string ConvertFromTimestampToString(double timestamp)
        {
            DateTime origin = new DateTime(1970, 1, 1, 0, 0, 0, 0);
            return string.Format("{0:yyyyMMdd}", origin.AddSeconds(timestamp));

        }



        public static DateTime ConvertFromTimestamp(double timestamp)
        {
            DateTime origin = new DateTime(1970, 1, 1, 0, 0, 0, 0);
            return origin.AddSeconds(timestamp);
        }

        public static double ConvertToTimestamp(this DateTime date)
        {
            DateTime origin = new DateTime(1970, 1, 1, 0, 0, 0, 0);
            TimeSpan diff = date - origin;
            return Math.Floor(diff.TotalSeconds);
        }



    }

}
