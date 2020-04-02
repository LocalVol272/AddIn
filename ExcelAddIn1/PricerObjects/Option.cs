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
}