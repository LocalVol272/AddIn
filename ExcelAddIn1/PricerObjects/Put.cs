namespace ExcelAddIn1.PricerObjects
{
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