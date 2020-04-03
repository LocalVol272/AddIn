﻿using System;
using System.Collections.Generic;
using MathNet.Numerics.LinearAlgebra.Double;
using MathNet.Numerics.LinearAlgebra.Factorization;

//To add from NuGet : MathNet
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
        //Polynomial fitting
        public static double[] Polyfit(double[] x, double[] y, int degree)
        {
            var v = new DenseMatrix(x.Length, degree + 1);
            for (int i = 0; i < v.RowCount; i++)
                for (int j = 0; j <= degree; j++) v[i, j] = Math.Pow(x[i], j);
            var yv = new DenseVector(y).ToColumnMatrix();
            QR<double> qr = v.QR();
            var r = qr.R.SubMatrix(0, degree + 1, 0, degree + 1);
            var q = v.Multiply(r.Inverse());
            var p = r.Inverse().Multiply(q.TransposeThisAndMultiply(yv));
            return p.Column(0).ToArray();
        }
        
        //Getting 1st or 2nd derivative from Polyfit output at a given point x
        public static double PolynomialDerivative(double[] coefs , double x , int derivOrder)
        {
            double res = 0.0;
            if (derivOrder == 1)
            {
                res = 3 * coefs[3] * x * x + 2 * coefs[2]*x + coefs[1];
            }
            if (derivOrder == 2)
            {
                res = 6 * coefs[3]*x+2* coefs[2];
            }
            return res;
        }

        public static Dictionary<string, double[,]> Sensitivities(double[,] price, double[] listK, double[] listT)
        {

            int nrows = listK.Length;
            int ncols = listT.Length;

            var dK = new double[nrows, ncols];
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
                        Array.Resize(ref tabP_fixedT, i + 1);
                        Array.Resize(ref tempK, i + 1);
                        tabP_fixedT[i] = price[j, t];
                        tempK[i] = listK[j];
                        i++;

                    }

                }
                //Then build cubic spline projection :

                //CubicSpline splineP_fixedT = new CubicSpline(tempK, tabP_fixedT);

                double[] coefs_fixedT = Grid.Polyfit(tempK, tabP_fixedT,3);

                for (var k = 0; k < nrows; k++)
                {
                    // Then get Price for fixed k :
                    double[] tabP_fixedK = new double[] { };
                    double[] tempT = new double[] { };
                    i = 0;

                    for (int j = 0; j < ncols; j++)
                    {
                        if (price[k, j] > 0.0)
                        {
                            Array.Resize(ref tabP_fixedK, i + 1);
                            Array.Resize(ref tempT, i + 1);
                            tabP_fixedK[i] = price[k, j];
                            tempT[i] = listT[j];
                            i++;
                        }

                    }
                    //Then build cubic spline projection :
                    //CubicSpline splineP_fixedK = new CubicSpline(tempT, tabP_fixedK);
                    double[] coefs_fixedK = Grid.Polyfit(tempT, tabP_fixedK, 3);

                    //Finally collect sensitivities
                    //dT[k, t] = splineP_fixedK.SpotEstimateSlope(listT[t]);
                    //dK[k, t] = splineP_fixedT.SpotEstimateSlope(listK[k]);
                    //dK2[k, t] = splineP_fixedT.SpotEstimateSecondDeriv(listK[k]);
                    dT[k, t] = Grid.PolynomialDerivative(coefs_fixedK, listT[t],1);
                    dK[k, t] = Grid.PolynomialDerivative(coefs_fixedT, listK[k], 1);
                    dK2[k, t] = Grid.PolynomialDerivative(coefs_fixedK, listK[k], 2);

                }
            }
            var dict = new Dictionary<string, double[,]>();
            dict.Add("dK", dK);
            dict.Add("dT", dT);
            dict.Add("dK2", dK2);
            return dict;
        }

        public double[,] LocalVolatility(double[,] price, double[] listK, double[] listT, double r)
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
