using System.Collections.Generic;

namespace ExcelAddIn1.PricerObjects
{
    public static class ApiMapping
    {
        public static readonly Dictionary<string, string> Roots = new Dictionary<string, string>()
        {
            {"GetAllTickers", "https://sandbox.iexapis.com/stable/ref-data/region/{0}/symbols?token={1}"},
            {"GetOptions", "https://query1.finance.yahoo.com/v7/finance/options/{0}?date={1}"},
            {"GetLastPrice", "https://sandbox.iexapis.com/stable/stock/{0}/previous?token={1}"}
        };
    }
}