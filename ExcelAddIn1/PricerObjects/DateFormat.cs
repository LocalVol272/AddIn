using System;
using System.Runtime.InteropServices;

namespace ProjetVolSto.PricerObjects
{
    class Date : IEXDate
    {
        string day, month, year;
        public string Day { get => day; set => day = value; }
        public string Month { get => month; set => month = value; }
        public string Year { get => year; set => year = value; }

        public Date(int year, int month, [Optional] int day)
        {
            Year = year.ToString(); Month = month.ToString(); Day = day.ToString();
        }


        public Date(string date)
        {
            if (date.Length == 8)
            {
                Year = date.Substring(0, 4);
                Month = date.Substring(4, 2);
                Day = date.Substring(6, 2);
            }
            else if (date.Length == 6)
            {
                Year = date.Substring(0, 4);
                Month = date.Substring(4, 2);
            }
            else { throw new Exception(DataLoaderError.DateFormatError); }
        }



        public string Format()
        {
            if (Day is null)
            {
                return String.Format("{0}{1}", year, month);
            }
            else
            {
                return String.Format("{0}{1}{2}", year, month, day);
            }

        }




    }
}