using System.Collections.Generic;
using ExcelAddIn1.PricerObjects;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace UniTestPricerVolSto
{
    [TestClass]
    public class TestApiRequest
    {
        public static Dictionary<string, object> getConfig()
        {
            var TickerList = new List<string> {"AAPL"};
            var Params = new Dictionary<string, object>();
            var DateList = new List<string> {"20200918", "20200403", "20210618"};
            Params.Add("ProductType", "Option/Call");
            Params.Add("Tickers", TickerList);
            Params.Add("Dates", new List<string>());

            var Config = new Dictionary<string, object>();
            Config.Add("Token", "Tsk_bbe66f58b6d149f59a9af4eb83bfc7f5");
            Config.Add("Type", "GET");
            Config.Add("Params", Params);

            return Config;
        }


        [TestMethod]
        public void apiRequest()
        {
            var config = getConfig();
            ApiRequest request = new ApiRequest(config);
            Assert.IsInstanceOfType(request, typeof(ApiRequest));
            Assert.IsInstanceOfType(request.RequestContent, typeof(YahooRequest));
        }
    }
}