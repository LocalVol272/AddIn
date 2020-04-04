using ExcelAddIn1;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelTest
{
    [TestClass]
    public class VolatilyCalculationTest
    {
        private Parameters _details = new Parameters();

        [TestMethod]
        public void TestObjectSettings()
        {
            MockParameters();
            Parameters details2 = new Parameters();
            VolatilyCalculation.VolatilyMain(details2);

            Assert.AreEqual(_details.ToString(), details2.ToString());
        }

        [TestMethod]
        public void TestObjectSettingsWithNullValues()
        {
            MockParameters();
            _details.spot = 2100;
            Parameters details2 = new Parameters();
            VolatilyCalculation.VolatilyMain(details2);

            Assert.AreNotEqual(_details.spot, details2.spot);
        }

        [TestMethod]
        public void AssertEndOfFunction()
        {
            //doit nous retourner last row = mstrikes.lengths +2
            MockParameters();
            _details = MockParametersOut();
            VolatilyCalculation.VolatilyMain(_details);

            Assert.AreEqual(_details.lastRow, _details.mStrikes.Length + 2);
        }

        private void MockParameters()
        {
            _details.lastRow = 5;
            _details.spot = 210;
            _details.mOptionMarketPrice = new double[10, 15];
            _details.mStrikes = new double[10];
            _details.moneyness = 0;
            _details.r = 0.02;
            _details.strikes = new double[10];
            _details.tenors = new double[15];
        }

        private Parameters MockParametersOut()
        {
            Parameters outData = new Parameters();
            outData.lastRow = 12;
            outData.spot = 250;
            outData.mOptionMarketPrice = new double[10, 15];
            outData.mStrikes = new double[10];
            outData.moneyness = 0;
            outData.r = 0.02;
            outData.strikes = new double[10];
            outData.tenors = new double[15];
            return outData;
        }
    }
}