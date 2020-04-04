using System;
using System.Collections.Generic;
using System.Linq;
using Extreme.Mathematics;
using Extreme.Statistics;

namespace ExcelAddIn1.PricingCalculation
{
    public class MatrixDecomposition : Matrix<double>
    {
        public MatrixDecomposition(int n) : base(n)
        {
        }

        public MatrixDecomposition(int m, int n) : base(m, n)
        {
        }

        public MatrixDecomposition(int m, int n, double x) : base(m, n, x)
        {
        }

        public MatrixDecomposition(double[,] source) : base(source)
        {
        }

        public MatrixDecomposition(Matrix<double> A) : base(A)
        {
        }

        protected override Matrix<double> Clone()
        {
            return new MatrixDecomposition(this);
        }

        protected override Matrix<double> CreateMatrix(int m, int n)
        {
            return new MatrixDecomposition(m, n);
        }

        protected override double Negative(double x)
        {
            return -x;
        }

        protected override double Add(double x, double y)
        {
            return x + y;
        }

        protected override double Multiply(double x, double y)
        {
            return x * y;
        }

        protected override double Sqrt(double x)
        {
            return Math.Sqrt(x);
        }

        public Dictionary<string, Matrix<double>> CholeskyDecomposition()
        {
            if (nbRows != nbCols)
                throw new Exception(
                    $"Cannot compute Cholesky decomposition, matrix is not square:\n\tA : {nbRows}x{nbCols}");
            var res = CreateMatrix(nbRows, nbCols);
            double dscnt;
            for (var i = 0; i < nbRows; i++)
            for (var j = 0; j < nbCols; j++)
            {
                dscnt = this[i, j];
                for (var h = 0; h < i; h++) dscnt -= res[i, h] * res[j, h];
                if (i == j)
                    res[i, j] = Sqrt(dscnt);
                else
                    res[j, i] = dscnt / res[i, i];
            }

            var dict = new Dictionary<string, Matrix<double>>();
            dict.Add("L", res);
            dict.Add("Lt", res.Transpose());
            return dict;
        }

        public Dictionary<string, Matrix<double>> LUDecomposition()
        {
            if (nbRows != nbCols)
                throw new Exception($"Cannot compute LU decomposition, matrix is not square:\n\tA : {nbRows}x{nbCols}");
            var L = CreateMatrix(nbRows, nbCols);
            var U = CreateMatrix(nbRows, nbCols);
            double dummy = 0;
            for (var i = 0; i < nbRows; i++) L[i, i] = 1;
            for (var i = 0; i < nbRows; i++)
            for (var j = 0; j < nbCols; j++)
            {
                dummy = this[i, j];
                if (i <= j)
                {
                    for (var h = 0; h < i; h++) dummy -= U[h, j] * L[i, h];
                    U[i, j] = dummy;
                }
                else
                {
                    for (var h = 0; h < j; h++) dummy -= U[h, j] * L[i, h];
                    L[i, j] = dummy / U[j, j];
                }
            }

            var dict = new Dictionary<string, Matrix<double>>();
            dict.Add("L", L);
            dict.Add("U", U);
            return dict;
        }

        public double[] ThomasAlgorithm(double[] r)
        {
            if (nbRows != nbCols)
                throw new Exception($"Cannot compute Thomas algoithm, matrix is not square:\n\tA : {nbRows}x{nbCols}");
            if (nbRows != r.Length)
                throw new Exception($"Cannot compute Thomas algoithm, dimension error :\n\tA : {nbCols}x{r.Length}");
            var lowerDiag = new double[nbCols - 1];
            var diag = new double[nbCols];
            var upperDiag = new double[nbCols - 1];
            var row_ = 0;
            for (var col = 0; col < nbCols - 1; col++)
            {
                diag[col] = this[row_, col];
                lowerDiag[col] = this[row_ + 1, col];
                upperDiag[col] = this[row_, col + 1];
                row_++;
            }

            diag[nbCols - 1] = this[nbRows - 1, nbCols - 1];
            var rStar_ = new double[nbCols];
            var upperDiagStar_ = new double[nbCols];
            rStar_[0] = r[0] / diag[0];
            upperDiagStar_[0] = upperDiag[0] / diag[0];
            for (var i = 1; i < nbCols; i++)
            {
                var m = 1 / (diag[i] - lowerDiag[i - 1] * upperDiagStar_[i - 1]);
                upperDiagStar_[i] = upperDiag[i - 1] * m;
                rStar_[i] = (r[i] - lowerDiag[i - 1] * rStar_[i - 1]) * m;
            }

            return rStar_;
        }

        public static List<double> PolynomialRegression(double[] x, double[] y, int degree)
        {
            Vector<double> x_regress = Vector.Create(x);
            Vector<double> y_regress = Vector.Create(y);

            var model = new PolynomialRegressionModel(x_regress, y_regress, 2);
            model.Fit();

            var coeff = model.Parameters.Select(param => param.Value).ToList();
            coeff.Reverse();
            return coeff;
        }
    }
}