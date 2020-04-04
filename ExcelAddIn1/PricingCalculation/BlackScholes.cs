using System;

namespace ExcelAddIn1.PricingCalculation
{
    public class BlackScholes
    {
        private const double pi = Math.PI;
        private readonly double _K;
        private readonly double _r;
        private readonly double _S;
        private readonly double _sigma;
        private readonly double _T;
        private readonly string _type;

        public BlackScholes(double S, double K, double T, double r, double sigma, string type)
        {
            _S = S;
            _K = K;
            _T = T;
            _r = r;
            _sigma = sigma;
            _type = type;
        }

        public double Compute()
        {
            var d1 = 0.0;
            var d2 = 0.0;
            var price = 0.0;

            d1 = (Math.Log(_S / _K) + (_r + _sigma * _sigma / 2.0) * _T) / (_sigma * Math.Sqrt(_T));
            d2 = d1 - _sigma * Math.Sqrt(_T);
            if (_type == "Call")
                price = _S * CND(d1) - _K * Math.Exp(-_r * _T) * CND(d2);
            else if (_type == "Put") price = _K * Math.Exp(-_r * _T) * CND(-d2) - _S * CND(-d1);
            return price;
        }

        private double CND(double x)
        {
            var L = 0.0;
            var K = 0.0;
            var res = 0.0;
            const double a1 = 0.31938153;
            const double a2 = -0.356563782;
            const double a3 = 1.781477937;
            const double a4 = -1.821255978;
            const double a5 = 1.330274429;
            L = Math.Abs(x);
            K = 1.0 / (1.0 + 0.2316419 * L);
            res = 1.0 - 1.0 / Math.Sqrt(2 * pi) * Math.Exp(-L * L / 2.0) *
                  (a1 * K + a2 * K * K + a3 * Math.Pow(K, 3.0) + a4 * Math.Pow(K, 4.0) + a5 * Math.Pow(K, 5.0));

            if (x < 0)
                return 1.0 - res;
            return res;
        }
    }
}