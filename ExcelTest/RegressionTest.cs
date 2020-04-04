using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelTest
{
    [TestClass]
    public class RegressionTest
    {
        [TestMethod]
        public void TestRegression()
        {
            double[,] price =
                {{10.0, 7.0, 5.0, 1.0}, {11.0, 7.0, 5.0, 1.0}, {12.0, 7.0, 5.0, 1.0}, {13.0, 7.0, 5.0, 1.0}};
            double[] listK = {0.0, 1.0, 2.0, 3.0};
            double[] listT = {0.0, 1.0, 2.0, 3.0};
            double[] tabP_fixedT = { };
            double[] tabP_fixedK = { };
            double[] tempK = { };
            double[] tempT = { };
            var i = 0;

            for (var j = 0; j < 4; j++)
                if (price[0, j] > 0.0)
                {
                    Array.Resize(ref tabP_fixedT, i + 1);
                    Array.Resize(ref tempK, i + 1);
                    tabP_fixedT[i] = price[0, j];
                    tempK[i] = listK[j];
                    i++;
                }

            i = 0;
            for (var j = 0; j < 4; j++)
                if (price[j, 0] > 0.0)
                {
                    Array.Resize(ref tabP_fixedK, i + 1);
                    Array.Resize(ref tempT, i + 1);
                    tabP_fixedK[i] = price[j, 0];
                    tempT[i] = listT[j];
                    i++;
                }

            for (var j = 0; j < tabP_fixedK.Length; j++)
            {
                Console.WriteLine(tempT[j]);
                Console.WriteLine(tabP_fixedK[j]);
            }

            Assert.AreEqual(tempT[0], 0);
        }
    }
}