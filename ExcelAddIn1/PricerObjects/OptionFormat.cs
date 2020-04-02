using System.Collections.Generic;

namespace ExcelAddIn1.PricerObjects
{
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
}