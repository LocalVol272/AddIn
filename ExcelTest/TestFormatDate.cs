using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace UniTestPricerVolSto
{
    [TestClass]
    public class TestFormatDate
    {
        [DataTestMethod]
        [DataRow("20201010")]
        public void formatDate(string date)
        {
            var test_date = new DateTime(2020, 10, 10);
            Assert.AreEqual(test_date.Year, 2020);
            Assert.AreEqual(test_date.Month, 10);
            Assert.AreEqual(test_date.Day, 10);
        }

        [TestMethod]
        [DataRow(2020, 10, 10)]
        public void formatDate(int year, int month, int day)
        {
            var test_date = new DateTime(year, month, day);
            var timestamp_date = Convert.ToDouble(test_date.TimeOfDay);
            Assert.AreEqual(timestamp_date, 1602288000.0);
        }
    }
}