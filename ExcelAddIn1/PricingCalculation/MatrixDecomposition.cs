using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LocalVolatility
{
    public class MatrixDecomposition : Matrix<double>
    {
        public MatrixDecomposition(int n) : base(n) { }
        public MatrixDecomposition(int m, int n) : base(m, n) { }
        public MatrixDecomposition(int m, int n, double x) : base(m, n, x) { }
        public MatrixDecomposition(double[,] source) : base(source) { }
        public MatrixDecomposition(Matrix<double> A) : base(A) { }

        protected override Matrix<double> Clone() => new MatrixDecomposition(this);
        protected override Matrix<double> CreateMatrix(int m, int n) => new MatrixDecomposition(m, n);
        protected override double Negative(double x) => -x;
        protected override double Add(double x, double y) => x + y;
        protected override double Multiply(double x, double y) => x * y;
        protected override double Sqrt(double x) => Math.Sqrt(x);

        public Dictionary<string, Matrix<double>> CholeskyDecomposition()
        {
            if (nbRows != nbCols)
            {
                throw new Exception($"Cannot compute Cholesky decomposition, matrix is not square:\n\tA : {nbRows}x{nbCols}");
            }
            Matrix<double> res = CreateMatrix(nbRows, nbCols);
            double dscnt;
            for (int i = 0; i < nbRows; i++)
            {
                for (int j = 0; j < nbCols; j++)
                {
                    dscnt = this[i, j];
                    for (int h = 0; h < i; h++)
                    {
                        dscnt -= res[i, h] * res[j, h];
                    }
                    if (i == j)
                    {
                        res[i, j] = Sqrt(dscnt);
                    }
                    else
                    {
                        res[j, i] = dscnt / res[i, i];
                    }
                }
            }
            Dictionary<string, Matrix<double>> dict = new Dictionary<string, Matrix<double>>();
            dict.Add("L", res);
            dict.Add("Lt", res.Transpose());
            return dict;
        }
        public Dictionary<string, Matrix<double>> LUDecomposition()
        {
            if (nbRows != nbCols)
            {
                throw new Exception($"Cannot compute LU decomposition, matrix is not square:\n\tA : {nbRows}x{nbCols}");
            }
            Matrix<double> L = CreateMatrix(nbRows, nbCols);
            Matrix<double> U = CreateMatrix(nbRows, nbCols);
            double dummy = 0;
            for (int i = 0; i < nbRows; i++)
            {
                L[i, i] = 1;
            }
            for (int i = 0; i < nbRows; i++)
            {
                for (int j = 0; j < nbCols; j++)
                {
                    dummy = this[i, j];
                    if (i <= j)
                    {
                        for (int h = 0; h < i; h++)
                        {
                            dummy -= U[h, j] * L[i, h];
                        }
                        U[i, j] = dummy;
                    }
                    else
                    {
                        for (int h = 0; h < j; h++)
                        {
                            dummy -= U[h, j] * L[i, h];
                        }
                        L[i, j] = dummy / U[j, j];
                    }
                }
            }
            Dictionary<string, Matrix<double>> dict = new Dictionary<string, Matrix<double>>();
            dict.Add("L", L);
            dict.Add("U", U);
            return dict;
        }
        public double[] ThomasAlgorithm(double[] r)
        {
            if (nbRows != nbCols)
            {
                throw new Exception($"Cannot compute Thomas algoithm, matrix is not square:\n\tA : {nbRows}x{nbCols}");
            }
            if (nbRows != r.Length)
            {
                throw new Exception($"Cannot compute Thomas algoithm, dimension error :\n\tA : {nbCols}x{r.Length}");
            }
            double[] lowerDiag = new double[nbCols - 1];
            double[] diag = new double[nbCols];
            double[] upperDiag = new double[nbCols - 1];
            int row_ = 0;
            for (int col = 0; col < nbCols - 1; col++)
            {
                diag[col] = this[row_, col];
                lowerDiag[col] = this[row_ + 1, col];
                upperDiag[col] = this[row_, col + 1];
                row_++;
            }
            diag[nbCols - 1] = this[nbRows - 1, nbCols - 1];
            double[] rStar_ = new double[nbCols];
            double[] upperDiagStar_ = new double[nbCols];
            rStar_[0] = r[0] / diag[0];
            upperDiagStar_[0] = upperDiag[0] / diag[0];
            for (int i = 1; i < nbCols; i++)
            {
                double m = 1 / (diag[i] - lowerDiag[i - 1] * upperDiagStar_[i - 1]);
                upperDiagStar_[i] = upperDiag[i - 1] * m;
                rStar_[i] = (r[i] - lowerDiag[i - 1] * rStar_[i - 1]) * m;
            }
            return rStar_;
        }
    }

}
