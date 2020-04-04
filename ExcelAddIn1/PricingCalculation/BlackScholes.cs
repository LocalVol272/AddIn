using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAddIn1.PricingCalculation
{
    public class BlackScholes
    {
        private double _S;
        private double _K;
        private double _T;
        private double _r;
        private double _sigma;
        private string _type;
        private const double pi = Math.PI;

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
            double d1 = 0.0;
            double d2 = 0.0;
            double price = 0.0;

            d1 = (Math.Log(_S / _K) + (_r + _sigma * _sigma / 2.0) * _T) / (_sigma * Math.Sqrt(_T));
            d2 = d1 - _sigma * Math.Sqrt(_T);
            if (_type == "Call")
            {
                price = _S * CND(d1) - _K * Math.Exp(-_r * _T) * CND(d2);
            }
            else if (_type == "Put")
            {
                price = _K * Math.Exp(-_r *_T) * CND(-d2) - _S * CND(-d1);
            }
            return price;
        }
        private double CND(double x)
        {
            double L = 0.0;
            double K = 0.0;
            double res = 0.0;
            const double a1 = 0.31938153;
            const double a2 = -0.356563782;
            const double a3 = 1.781477937;
            const double a4 = -1.821255978;
            const double a5 = 1.330274429;
            L = Math.Abs(x);
            K = 1.0 / (1.0 + 0.2316419 * L);
            res = 1.0 - 1.0 / Math.Sqrt(2 * pi) * Math.Exp(-L * L / 2.0) * (a1 * K + a2 * K * K + a3 * Math.Pow(K, 3.0) + a4 * Math.Pow(K, 4.0) + a5 * Math.Pow(K, 5.0));

            if (x < 0)
            {
                return 1.0 - res;
            }
            else
            {
                return res;
            }
        }
    }
}
