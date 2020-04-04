using System;

namespace ExcelAddIn1.PricingCalculation
{
    public class CubicSpline
    {
        //last checked index, for estimation
        private int _previousI;

        // n-1 spline coefficients for n points
        private double[] a;
        private double[] b;

        // Initial x and y 
        private double[] xInit;
        private double[] yInit;

        // Default constructor
        public CubicSpline()
        {
        }

        // Construct and call Fit.
        //  x  coordinates to fit
        // y coordinates to fit
        // startSlope Optional slope constraint for the first point. Single.NaN means no constraint
        // endSlope Optional slope constraint for the final point. Single.NaN means no constraint
        // debug Turn on console output. Default is false
        public CubicSpline(double[] x, double[] y, double startSlope = double.NaN, double endSlope = double.NaN,
            bool debug = false)
        {
            Fit(x, y, startSlope, endSlope, debug);
        }

        // Throws if Fit has not been called.
        private void IsFitted()
        {
            if (a == null) throw new Exception("Fit before estimation !");
        }


        // Find the i and i+1 such that xInit[i] < x < xInit[i+1] :
        private int _NextI(double x)
        {
            while (_previousI < xInit.Length - 2 && x > xInit[_previousI + 1]) _previousI++;

            return _previousI;
        }

        // Estimate y(x) value using spline
        // j which spline to use

        private double EstimateSpline(double x, int j, bool debug = false)
        {
            var dx = xInit[j + 1] - xInit[j];
            var t = (x - xInit[j]) / dx;
            var y = (1 - t) * yInit[j] + t * yInit[j + 1] + t * (1 - t) * (a[j] * (1 - t) + b[j] * t); // equation 9
            if (debug) Console.WriteLine("xs = {0}, j = {1}, t = {2}", x, j, t);
            return y;
        }


        // Fit x,y and then eval at points xs and return the corresponding y's.
        // This does the "natural spline" style for ends.
        // This can extrapolate off the ends of the splines.
        // You must provide points in X sort order.
        // x X coordinates to fit.
        // y Y coordinates to fit
        // xs X coordinates to evaluate the fitted curve at
        // startSlope Optional slope constraint for the first point. Single.NaN means no constraint
        // endSlope Optional slope constraint for the final point. Single.NaN means no constraint
        // debug Turn on console output. Default is false.
        // returns the computed y values for each xs.
        public double[] FitAndEstimate(double[] x, double[] y, double[] xs, double startSlope = double.NaN,
            double endSlope = double.NaN, bool debug = false)
        {
            Fit(x, y, startSlope, endSlope, debug);
            return Estimate(xs, debug);
        }

        public void Fit(double[] x, double[] y, double startSlope = double.NaN, double endSlope = double.NaN,
            bool debug = false)
        {
            /*if (Single.IsInfinity(startSlope) || Single.IsInfinity(endSlope))
            {
                throw new Exception("startSlope and endSlope cannot be infinity.");
            }*/
            //save initial values (for future estimates)
            xInit = x;
            yInit = y;

            var n = x.Length;
            var r = new double[n];

            //tridiagonal matrix
            var m = new double[n, n];
            for (var i = 0; i < n; i++)
            for (var j = 0; j < n; j++)
                m[j, i] = 0;

            double dx1, dx2, dy1, dy2;

            //Fill-in initial values of tridiagonal matrix for diag and upper diag :
            if (double.IsNaN(startSlope))
            {
                dx1 = x[1] - x[0];

                m[0, 1] = 1.0 / dx1;
                m[0, 0] = 2.0 * m[0, 1];
                r[0] = 3 * (y[1] - y[0]) / (dx1 * dx1);
            }
            else
            {
                m[0, 0] = 1;
                r[0] = startSlope;
            }

            // Fill-in middle values for all rows :
            for (var i = 1; i < n - 1; i++)
            {
                dx1 = x[i] - x[i - 1];
                dx2 = x[i + 1] - x[i];

                m[i, i - 1] = 1.0 / dx1;
                m[i, i + 1] = 1.0 / dx2;
                m[i, i] = 2.0 * (m[i, i - 1] + m[i, i + 1]);

                dy1 = y[i] - y[i - 1];
                dy2 = y[i + 1] - y[i];

                r[i] = 3 * (dy1 / (dx1 * dx1) + dy2 / (dx2 * dx2));
            }


            //Fill-in ending values of diag and lower diag :
            if (double.IsNaN(endSlope))
            {
                dx1 = x[n - 1] - x[n - 2];
                dy1 = y[n - 1] - y[n - 2];

                m[n - 1, n - 2] = 1.0 / dx1;
                m[n - 1, n - 1] = 2.0 * m[n - 1, n - 2];

                r[n - 1] = 3 * (dy1 / (dx1 * dx1));
            }
            else
            {
                m[n - 1, n - 1] = 1;
                r[n - 1] = endSlope;
            }


            //==================================================================
            //==================================================================
            //==================================================================
            //==================================================================
            //Thomas algoithm===================================================
            var md = new MatrixDecomposition(m);
            if (debug) Console.WriteLine("Matrix:\n{0}", md);
            var k = md.ThomasAlgorithm(r);
            //==================================================================
            // we want k, the solution to the matrix
            //==================================================================
            //==================================================================

            // slpine coefs a and b
            a = new double[n - 1];
            b = new double[n - 1];

            for (var i = 1; i < n; i++)
            {
                dx1 = x[i] - x[i - 1];
                dy1 = y[i] - y[i - 1];
                a[i - 1] = k[i - 1] * dx1 - dy1;
                b[i - 1] = -k[i] * dx1 + dy1;
            }
        }


        // The following isn't multithreading proof (yet) !! (watch out Yugo)
        //problem is with resetting _previousI
        public double[] Estimate(double[] x, bool debug = false)
        {
            IsFitted();

            var n = x.Length;
            var retY = new double[n];
            _previousI = 0; //for multiple estimations, set to 0 at each eval

            for (var i = 0; i < n; i++)
            {
                // get the right spline to use :
                var j = _NextI(x[i]);

                // Estimate with j spline
                retY[i] = EstimateSpline(x[i], j, debug);
            }

            return retY;
        }


        public double[] EstimateSlope(double[] x, bool debug = false)
        {
            //first check if fitted :
            IsFitted();

            var n = x.Length;
            var retSlopes = new double[n];

            _previousI = 0; //for multiple estimations, set to 0 at each eval

            for (var i = 0; i < n; i++)
            {
                // Find which spline can be used to compute this x (by simultaneous traverse)
                var j = _NextI(x[i]);

                // Estimate at j spline
                var dx = xInit[j + 1] - xInit[j];
                var dy = yInit[j + 1] - yInit[j];
                var t = (x[i] - xInit[j]) / dx;


                retSlopes[i] = dy / dx + (1 - 2 * t) * (a[j] * (1 - t) + b[j] * t) / dx +
                               t * (1 - t) * (b[j] - a[j]) / dx;

                if (debug) Console.WriteLine("[{0}]: xs = {1}, j = {2}, t = {3}", i, x[i], j, t);
            }

            return retSlopes;
        }

        public double SpotEstimateSlope(double x, bool debug = false)
        {
            //first check if fitted :
            IsFitted();


            double retSlope;

            _previousI = 0; //for multiple estimations, set to 0 at each eval


            // Find which spline can be used to compute this x (by simultaneous traverse)
            var j = _NextI(x);

            // Estimate at j spline
            var dx = xInit[j + 1] - xInit[j];
            var dy = yInit[j + 1] - yInit[j];
            var t = (x - xInit[j]) / dx;


            retSlope = dy / dx + (1 - 2 * t) * (a[j] * (1 - t) + b[j] * t) / dx + t * (1 - t) * (b[j] - a[j]) / dx;

            if (debug) Console.WriteLine("[{0}]: xs = {1}, j = {2}, t = {3}", x, j, t);


            return retSlope;
        }


        public double SpotEstimateSecondDeriv(double x, bool debug = false)
        {
            //first check if fitted :
            IsFitted();


            double retSecDer;

            _previousI = 0; //for multiple estimations, set to 0 at each eval


            // Find which spline can be used to compute this x (by simultaneous traverse)
            var j = _NextI(x);

            // Estimate at j spline
            var dx = xInit[j + 1] - xInit[j];
            var t = (x - xInit[j]) / dx;

            retSecDer = 2 * (b[j] - 2 * a[j] + (a[j] - b[j]) * 3 * t) / (dx * dx);

            if (debug) Console.WriteLine("[{0}]: xs = {1}, j = {2}, t = {3}", x, j, t);


            return retSecDer;
        }
    }
}