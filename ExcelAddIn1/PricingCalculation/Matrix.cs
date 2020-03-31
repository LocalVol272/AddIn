using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LocalVolatility
{
    public abstract class Matrix<T>
    {
        private readonly T[,] data;
        public int nbRows { get; }
        public int nbCols { get; }
        public T this[int i, int j]
        {
            get => data[i, j];
            set => data[i, j] = value;
        }
        protected Matrix(int n) : this(n, n) { }
        protected Matrix(int m, int n) : this(m, n, default(T)) { }
        protected Matrix(Matrix<T> A) : this(A.data) { }
        protected Matrix(T[,] source)
        {
            nbRows = source.GetLength(0);
            nbCols = source.GetLength(1);
            this.data = new T[nbRows, nbCols];
            this.SetValues((i, j) => source[i, j]);
        }

        protected Matrix(int m, int n, T x)
        {
            if (m < 1)
            {
                throw new Exception($"Cannot create matrix, M should be greater or equal to 1. \n\tM = {m}");
            }
            if (n < 1)
            {
                throw new Exception($"Cannot create matrix, N should be greater or equal to 1. \n\tN = {n}");
            }
            nbRows = m;
            nbCols = n;
            data = new T[m, n];
            Fill(x);
        }
        public void Fill(T x) => SetValues(x);
        public static Matrix<T> operator -(Matrix<T> a) => a.Negative();
        public static Matrix<T> operator ~(Matrix<T> a) => a.Transpose();
        public static Matrix<T> operator +(Matrix<T> a, Matrix<T> b) => a.Add(b);
        public static Matrix<T> operator -(Matrix<T> a, Matrix<T> b) => a + (-b);
        public static Matrix<T> operator *(Matrix<T> a, Matrix<T> b) => a.Multiply(b);
        public static Matrix<T> operator +(Matrix<T> a, T x) => a.Add(x);
        public static Matrix<T> operator -(Matrix<T> a, T x) => a.Substract(x);
        public static Matrix<T> operator *(Matrix<T> a, T x) => a.MultiplyBy(x);
        public T Trace()
        {
            if (nbRows != nbCols)
            {
                throw new Exception($"Cannot compute trace, matrix is not square:\n\tA : {nbRows}x{nbCols}");
            }
            T res = default(T);
            for (int i = 0; i < nbRows; i++)
            {
                res = Add(res, this[i, i]);
            }
            return res;
        }
        public Matrix<T> Transpose() => Clone().SetValues((_, i, j) => this[j, i]);
        public Matrix<T> Diagonal()
        {
            if (nbRows != nbCols)
            {
                throw new Exception($"Cannot compute diagonal, matrix is not square:\n\tA : {nbRows}x{nbCols}");
            }
            Matrix<T> res = CreateMatrix(nbRows, nbCols);
            for (int i = 0; i < nbRows; i++)
            {
                res[i, i] = this[i, i];
            }
            return res;
        }
        public Matrix<T> Map(Func<T, T> func) => CreateMatrix(nbRows, nbCols).SetValues(func);
        public T[,] ToArray()
        {
            T[,] arr = new T[nbRows, nbCols];
            for (int i = 0; i < nbRows; i++)
            {
                for (int j = 0; j < nbCols; j++)
                {
                    arr[i, j] = this[i, j];
                }
            }
            return arr;
        }
        public override string ToString()
        {
            var sb = new StringBuilder();
            for (int i = 0; i < nbRows; i++)
            {
                sb.Append("[ ");
                for (int j = 0; j < nbCols; j++)
                {
                    string nb = AsString(this[i, j]);
                    if (j != nbCols - 1)
                    {
                        sb.Append($"{nb} ");
                    }
                    else
                    {
                        sb.Append($"{nb} ]{Environment.NewLine}");
                    }
                }
            }
            return sb.ToString();
        }
        protected abstract Matrix<T> Clone();
        protected abstract Matrix<T> CreateMatrix(int m, int n);
        private Matrix<T> CreateSameSizeMatrix() => CreateMatrix(nbRows, nbCols);
        protected abstract T Negative(T x);
        protected abstract T Add(T x, T y);
        protected abstract T Multiply(T x, T y);
        protected abstract T Sqrt(T x);
        protected virtual string AsString(T x) => x.ToString();
        private T Substract(T x, T y) => Add(x, Negative(y));
        private Matrix<T> Multiply(Matrix<T> B)
        {
            if (nbCols != B.nbRows)
            {
                throw new Exception($"Cannot compute diagonal, matrix dimension do not match:\n\tA : {nbRows}x{nbCols}\n\tB : {B.nbRows}x{B.nbCols}");
            }
            Matrix<T> res = CreateMatrix(nbRows, B.nbCols);
            res.SetValues((current, i, j) =>
            {
                var sum = default(T);
                for (int k = 0; k < nbCols; k++)
                {
                    sum = Add(sum, Multiply(this[i, k], B[k, j]));
                }
                return sum;
            });
            return res;
        }
        private Matrix<T> Negative() => Clone().SetValues(Negative);
        private Matrix<T> Add(Matrix<T> B)
        {
            if (nbRows != B.nbRows || nbCols != B.nbCols)
            {
                throw new Exception($"Cannot compute sum, dimensions do not match :\ntA : {nbRows}x{nbCols}\n\tB : {B.nbRows}x{B.nbCols}");
            }
            return Clone().SetValues((current, i, j) => Add(current, B[i, j]));
        }
        private Matrix<T> Add(T x) => CreateSameSizeMatrix().SetValues((i, j) => Add(this[i, j], x));
        private Matrix<T> Substract(T x) => Add(Negative(x));
        private Matrix<T> MultiplyBy(T x) => CreateSameSizeMatrix().SetValues((i, j) => Multiply(this[i, j], x));
        protected Matrix<T> SetValues(Func<T, int, int, T> newValueSeletor)
        {
            for (int i = 0; i < nbRows; i++)
            {
                for (int j = 0; j < nbCols; j++)
                {
                    this[i, j] = newValueSeletor(this[i, j], i, j);
                }
            }
            return this;
        }
        protected Matrix<T> SetValues(T newValue) => SetValues((_, __, ___) => newValue);
        protected Matrix<T> SetValues(Func<T, T> newValueSelector) => SetValues((current, _, __) => newValueSelector(current));
        protected Matrix<T> SetValues(Func<int, int, T> newValueSelector) => SetValues((_, i, j) => newValueSelector(i, j));

    }

}
