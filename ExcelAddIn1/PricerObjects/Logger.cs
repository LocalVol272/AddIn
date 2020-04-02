using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.PricerObjects
{
    public static class AvailableData
    {
        public static List<string> GetPath()
        {
            var fileName = "\\ExcelAddIn1\\TICKER.txt";

            string projectDir =Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\.."));
            var fileDir = projectDir + fileName;
            List<string> ticker = new List<string>();

            string[] lines = System.IO.File.ReadAllLines(fileDir);

            foreach (string line in lines)
            {
                ticker.Add(line);
            }

            return ticker;

        }

        public static List<string> GetTicker()
        {
            return GetPath();

        }
    }
}

