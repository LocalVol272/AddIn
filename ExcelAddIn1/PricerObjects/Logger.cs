using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelAddIn1.PricerObjects
{
    public static class AvailableData
    {
        public static List<string> GetTicker()
        {
            var fileName = "\\ExcelAddIn1\\TICKER.txt";
            var projectDir = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\.."));
            var fileDir = projectDir + fileName;
            var ticker = new List<string>();

            var lines = File.ReadAllLines(fileDir);

            foreach (var line in lines) ticker.Add(line);

            return ticker;
        }
    }
}