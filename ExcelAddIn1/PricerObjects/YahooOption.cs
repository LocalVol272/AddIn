using System.Collections.Generic;

namespace ExcelAddIn1.PricerObjects
{
    public class YahooOption
    {
        public double expirationDate { get; set; }
        public List<Call> calls { get; set; }
        public List<Put> puts { get; set; }
    }
}