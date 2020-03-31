using System.Collections.Generic;

namespace ProjetVolSto.Struct
{
    public class Option
    {
        public string symbol { get; set; }
        public string expirationDate { get; set; }
        public string strikePrice { get; set; }
        public string closingPrice { get; set; }
        public string type { get; set; }
        public string bid { get; set; }
        public string ask { get; set; }
    }

    public static class ApiMapping
    {
        public static readonly Dictionary<string, string> Roots = new Dictionary<string, string>()
        {
            {"GetAllTickers", "https://sandbox.iexapis.com/stable/ref-data/region/{0}/symbols?token={1}"},
            {"GetOptions", "https://sandbox.iexapis.com/stable/stock/{0}/options/{1}/{2}?token={3}"}
        };
    }
}