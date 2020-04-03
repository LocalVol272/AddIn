using System;
using System.Collections.Generic;

namespace ExcelAddIn1.PricingCalculation
{
    class Grid
    {
        private readonly double[,] prices;

        public Grid(double[,] source, double[] tenors, double[] strikes)
        {
            nbRows = source.GetLength(0);
            nbCols = source.GetLength(1);
            //Check if it contains null values and apply cubic spline if its the case
            prices = source;
            this.tenors = tenors;
            this.strikes = strikes;
            if (nbCols != tenors.Length)
                throw new Exception($"Cannot build Grid, dimension error :\n\tA : {nbCols}x{tenors.Length}");
            if (nbRows != strikes.Length)
                throw new Exception($"Cannot build Grid, dimension error :\n\tA : {nbRows}x{strikes.Length}");
        }

        public int nbRows { get; }
        public int nbCols { get; }
        public double[] tenors { get; }
        public double[] strikes { get; }

        public double this[int i, int j]
        {
            get => prices[i, j];
            set => prices[i, j] = value;
        }

        public static Dictionary<string, double[,]> Sensitivities(double[,] price, double[] listK, double[] listT)
        {
     
            int nrows = listK.Length;
            int ncols = listT.Length;

            var dK = new double[nrows,ncols];
            var dT = new double[nrows, ncols];
            var dK2 = new double[nrows, ncols];

            for (var t = 0; t < ncols; t++)
            {
                // t is fixed, get spline interpolation of (K,Price):
                //First get and Price[i]:
                double[] tabP_fixedT = new double[] { };
                double[] tempK = new double[] { };
                int i = 0;

                for (int j = 0; j < nrows; j++)
                {
                    if (price[j, t] > 0.0)
                    {
                        tabP_fixedT[i] = price[j, t];
                        tempK[i] = listK[j];
                        i++;

                    }

                }
                //Then build cubic spline projection :

                CubicSpline splineP_fixedT = new CubicSpline(tempK, tabP_fixedT);

                for (var k =0; k <nrows; k++)
                {
                    // Then get Price for fixed k :
                    double[] tabP_fixedK = new double[ncols];
                    double[] tempT = new double[] { };
                    i = 0;

                    for (int j = 0; j < ncols; j++)
                    {
                        if(price[k, j]>0.0)
                        {
                            tabP_fixedK[i] = price[k, j];
                            tempT[i] = listT[j];
                            i++;
                        }

                    }
                    //Then build cubic spline projection :
                    CubicSpline splineP_fixedK = new CubicSpline(tempT, tabP_fixedK);


                    //Finally collect sensitivities
                    dT[k, t] = splineP_fixedK.SpotEstimateSlope(listT[t]);
                    dK[k, t] = splineP_fixedT.SpotEstimateSlope(listK[k]);
                    dK2[k, t] = splineP_fixedT.SpotEstimateSecondDeriv(listK[k]);
                }
            }
            var dict = new Dictionary<string, double[,]>();
            dict.Add("dK", dK);
            dict.Add("dT", dT);
            dict.Add("dK2", dK2);
            return dict;
        }

        public static double[,] LocalVolatility(double[,] price, double[] listK, double[] listT, double r)
        {
            var sensiDict = Grid.Sensitivities(price, listK, listT);
            int nrows = listK.Length;
            int ncols = listT.Length;

            var locvol = new double[nrows, ncols];
            double[,] dT = sensiDict["dT"];
            double[,] dK = sensiDict["dK"];
            double[,] dK2 = sensiDict["dK2"];

            for (int i = 0; i < nrows; i++)
            {
                for (int j = 0; j < ncols; j++)
                {
                    Console.WriteLine("dT(" + i + "," + j + ") : " + dT[i, j]);
                    Console.WriteLine("dK(" + i + "," + j + ") : " + dK[i, j]);
                    Console.WriteLine("dk2(" + i + "," + j + ") : " + dK2[i, j]);
                    locvol[i, j] = Math.Sqrt((dT[i, j] + r * listK[i] * dK[i, j]-r*price[i,j]) /  0.5*Math.Pow(listK[i], 2) * dK2[i, j]);
                }
            }

                    
            return locvol;
        }

        public double[,] BSPD(double S, double r, double[,] VolLocale, string type)
        {
            double[,] price = new double[nbRows, nbCols];
            double K;
            double T;
            double sigma;
            for (int i = 0; i < nbRows; i++)
            {
                K = strikes[i];
                for (int j = 0; j < nbCols; j++)
                {
                    T = tenors[j];
                    sigma = VolLocale[i, j];
                    BlackScholes bs = new BlackScholes(S, K, T, r, sigma, type);
                    price[i, j] = bs.Compute();

                }
            }
            return price;
        }
    }
}
