using System;
using ExcelAddIn1.PricingCalculation;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelTest
{
    [TestClass]
    public class GridTest
    {
        [TestMethod]
        public void TryCubicSpline()
        {
            var x = new double[20];
            var y = new double[20];
            for (var i = 0; i < 20; i++)
            {
                x[i] = (i - 10.0) / 1.0;
                y[i] = -x[i] * x[i];
            }

            var x1 = new double[30];
            var y1 = new double[30];
            for (var i = 0; i < 30; i++)
            {
                x1[i] = (i - 10.0) / 1.0;
                y1[i] = -x1[i] * x1[i];
            }

            CubicSpline testSpline = new CubicSpline(x, y, double.NaN, double.NaN, true);
            CubicSpline testSpline1 = new CubicSpline(x1, y1, double.NaN, double.NaN, true);

            Assert.AreNotEqual(testSpline.Estimate(x, false), testSpline1.Estimate(y1, false));
        }

        [TestMethod]
        public void TryCubicSplineNumber2()
        {
            var x = new double[20];
            var y = new double[20];
            for (var i = 0; i < 20; i++)
            {
                x[i] = (i - 10.0) / 1.0;
                y[i] = -x[i] * x[i];
            }


            CubicSpline testSpline = new CubicSpline(x, y, double.NaN, double.NaN, true);


            var nNumb = 15;
            var evals = new double[nNumb];

            for (var i = 0; i < nNumb; i++) evals[i] = i / 10.0;

            double[] res = testSpline.Estimate(evals);
            double[] slopes = testSpline.EstimateSlope(evals);

            Console.Write("x  ");
            for (var i = 0; i < nNumb; i++) Console.Write(evals[i] + " ");
            Console.WriteLine();
            Console.Write("y  ");
            for (var i = 0; i < nNumb; i++) Console.Write(res[i] + " ");
            Console.WriteLine();
            Console.Write("D  ");
            for (var i = 0; i < nNumb; i++) Console.Write(slopes[i] + " ");
            Console.WriteLine();
            Console.Write("2D ");

            double secder = testSpline.SpotEstimateSecondDeriv(0.29);
            double secder1 = testSpline.SpotEstimateSecondDeriv(0.30);
            Console.WriteLine("MYSECDER : " + secder);

            Assert.AreNotEqual(secder, secder1);
        }
    }
}