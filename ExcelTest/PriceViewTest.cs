using System.Collections.Generic;
using ExcelAddIn1;
using ExcelAddIn1.PricerObjects;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelTest
{
    [TestClass]
    public class PriceViewTest
    {
        private readonly Parameters _details = new Parameters();

        [TestMethod]
        public void GetOptionsCallPutSameTicker()
        {
            var res = ReplicateGetOptions("AAPL", "Call");
            var res1 = ReplicateGetOptions("AAPL", "Put");
            Assert.AreNotEqual(res, res1);
        }

        private Dictionary<string, Dictionary<string, List<Option>>> ReplicateGetOptions(string ticker, string type)
        {
            //api virtuelle retourne dico
            if (ticker == "AAPL")
            {
                if (type == "Call")
                    return AAPLCALL();
                if (type == "Put") return AAPLPUT();
            }
            else if (ticker == "FB")
            {
                if (type == "Call")
                    return FBCALL();
                if (type == "Put") return FBPUT();
            }

            return new Dictionary<string, Dictionary<string, List<Option>>>();
        }

        [TestMethod]
        public void GetOptionsCallCallDifferentTicker()
        {
            var res = ReplicateGetOptions("AAPL", "Call");
            var res1 = ReplicateGetOptions("FB", "Call");
            Assert.AreNotEqual(res, res1);
        }

        [TestMethod]
        public void GetOptionsCallCallSameTicker()
        {
            var res = ReplicateGetOptions("AAPL", "Call");
            var res1 = ReplicateGetOptions("AAPL", "Call");
            Assert.AreEqual(res.Keys.ToString(), res1.Keys.ToString());
        }

        [TestMethod]
        public void GetOptionsPutPutSameTicker()
        {
            var res = ReplicateGetOptions("AAPL", "Put");
            var res1 = ReplicateGetOptions("AAPL", "Put");
            Assert.AreEqual(res.Keys.ToString(), res1.Keys.ToString());
        }

        private Dictionary<string, Dictionary<string, List<Option>>> AAPLCALL()
        {
            Dictionary<string, Dictionary<string, List<Option>>> aaplCaal =
                new Dictionary<string, Dictionary<string, List<Option>>>
                {
                    {"AAPL", new Dictionary<string, List<Option>>()},
                    {"Call", new Dictionary<string, List<Option>>()}
                };
            return aaplCaal;
        }

        private Dictionary<string, Dictionary<string, List<Option>>> AAPLPUT()
        {
            Dictionary<string, Dictionary<string, List<Option>>> aaplCaal =
                new Dictionary<string, Dictionary<string, List<Option>>>
                {
                    {"AAPL", new Dictionary<string, List<Option>>()},
                    {"Put", new Dictionary<string, List<Option>>()}
                };
            return aaplCaal;
        }

        private Dictionary<string, Dictionary<string, List<Option>>> FBCALL()
        {
            Dictionary<string, Dictionary<string, List<Option>>> aaplCaal =
                new Dictionary<string, Dictionary<string, List<Option>>>
                {
                    {"FB", new Dictionary<string, List<Option>>()},
                    {"Call", new Dictionary<string, List<Option>>()}
                };
            return aaplCaal;
        }

        private Dictionary<string, Dictionary<string, List<Option>>> FBPUT()
        {
            Dictionary<string, Dictionary<string, List<Option>>> aaplCaal =
                new Dictionary<string, Dictionary<string, List<Option>>>
                {
                    {"FB", new Dictionary<string, List<Option>>()},
                    {"Put", new Dictionary<string, List<Option>>()}
                };
            return aaplCaal;
        }
    }
}