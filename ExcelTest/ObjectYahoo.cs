using System;
using ExcelAddIn1.PricerObjects;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace UniTestPricerVolSto
{
    [TestClass]
    public class ObjectYahoo
    {
        [DataTestMethod]
        [DataRow("123456789", "123456789")]
        [DataRow("Key54684987984949494", "Key54684987984949494")]
        public void TestToken(string token_value, string result)
        {
            Token test = new Token(result);
            Token token1 = new Token(token_value);
            Assert.AreEqual(token1.value, test.value);
        }

        [DataTestMethod]
        [DataRow("AAPL", "20201209", "10", "12", "2", "1.98", "Put")]
        public void TestDataLoader(string Symbol, string ExpirationDate, string StrikePrice, string ClosingPrice,
            string Bid, string Ask, string Type)
        {
            var date = new DateTime(2020, 10, 12);
            Option option = new Option(Symbol, "1607554800", StrikePrice, ClosingPrice, Bid, Ask, Type);

            Assert.AreEqual(option.symbol, Symbol);
            Assert.AreEqual(option.expirationDate, ExpirationDate);
            Assert.AreEqual(option.strikePrice, StrikePrice);
            Assert.AreEqual(option.closingPrice, ClosingPrice);
            Assert.AreEqual(option.bid, Bid);
            Assert.AreEqual(option.ask, Ask);
            Assert.AreEqual(option.type, Type);
        }
    }
}