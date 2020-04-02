using System.Collections.Generic;

namespace ExcelAddIn1.PricerObjects
{
    public class YahooOptionChain
    {
        public string underlyingSymbol { get; set; }
        public List<double> expirationDates { get; set; }
        public List<double> strikes { get; set; }
        public Dictionary<string, string> quote { get; set; }
        public List<YahooOption> options { get; set; }
    }
}