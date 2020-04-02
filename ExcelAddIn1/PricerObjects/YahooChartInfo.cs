using System.Collections.Generic;

namespace ExcelAddIn1.PricerObjects
{
    public class YahooChartInfo
    {
        public List<double> timestamp { get; set; }
        public Indicators indicators { get; set; }
    }
}