using System.Collections.Generic;

namespace ExcelAddIn1.PricerObjects
{
    public class Option
    {
        public Option(string Symbol, string ExpirationDate, string StrikePrice, string ClosingPrice, string Bid,
            string Ask, string Type)
        {
            symbol = Symbol;
            expirationDate = double.Parse(ExpirationDate).ConvertFromTimestampToString();
            strikePrice = StrikePrice;
            closingPrice = ClosingPrice;
            bid = Bid;
            ask = Ask;
            type = Type;
        }

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
        public static readonly Dictionary<string, string> Roots = new Dictionary<string, string>
        {
            {"GetAllTickers", "https://sandbox.iexapis.com/stable/ref-data/region/{0}/symbols?token={1}"},
            {"GetOptions", "https://query1.finance.yahoo.com/v7/finance/options/{0}?date={1}"},
            {
                "GetLastPrice",
                "https://query1.finance.yahoo.com/v8/finance/chart/{0}?period1={1}&period2={2}&interval=1d&includePrePost=False&events=div%2Csplits"
            },
            {"GetAllAvailableMaturities", "https://query1.finance.yahoo.com/v7/finance/options/{0}"}
        };
    }


    public static class OptionFormat
    {
        public static string TypeCall = "Call";
        public static string TypePut = "Put";

        public static List<Option> ToListOption(this List<Call> calls)
        {
            var listOption = new List<Option>();
            calls.ForEach(x => listOption.Add(new Option(x.contractSymbol, x.expiration.ToString(), x.strike.ToString(),
                x.lastPrice.ToString(), x.bid.ToString(), x.ask.ToString(), TypeCall)));
            return listOption;
        }

        public static List<Option> ToListOption(this List<Put> calls)
        {
            var listOption = new List<Option>();
            calls.ForEach(x => listOption.Add(new Option(x.contractSymbol, x.expiration.ToString(), x.strike.ToString(),
                x.lastPrice.ToString(), x.bid.ToString(), x.ask.ToString(), TypePut)));
            return listOption;
        }
    }


    public class YahooOptionChain
    {
        public string underlyingSymbol { get; set; }
        public List<double> expirationDates { get; set; }
        public List<double> strikes { get; set; }
        public Dictionary<string, string> quote { get; set; }
        public List<YahooOption> options { get; set; }
    }


    public class YahooOption
    {
        public double expirationDate { get; set; }
        public List<Call> calls { get; set; }
        public List<Put> puts { get; set; }
    }


    public class Call
    {
        public string contractSymbol { get; set; }
        public double strike { get; set; }
        public double lastPrice { get; set; }
        public double ask { get; set; }
        public double bid { get; set; }

        public double expiration { get; set; }
    }

    public class Put
    {
        public string contractSymbol { get; set; }
        public double strike { get; set; }
        public double lastPrice { get; set; }
        public double ask { get; set; }
        public double bid { get; set; }

        public double expiration { get; set; }
    }
}