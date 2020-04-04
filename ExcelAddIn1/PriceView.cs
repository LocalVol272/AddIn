using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ExcelAddIn1.PricerObjects;

namespace ExcelAddIn1
{
    public static class PriceView
    {

        public static Dictionary<string, Dictionary<string, List<Option>>> GetOptions(string ticker,string combobox)
        {
            List<string> TickerList = new List<string>() { ticker };
            Dictionary<string, object> Params = new Dictionary<string, object>();
            List<string> DateList = new List<string>() { };

            Params.Add("ProductType", "Option/" + combobox);
            Params.Add("Tickers", TickerList);
            Params.Add("Dates", DateList);

            Dictionary<string, object> Config = new Dictionary<string, object>() { };
            Config.Add("Token", "Tsk_bbe66f58b6d149f59a9af4eb83bfc7f5");
            Config.Add("Type", "GET");
            Config.Add("Params", Params);

            Options test = new Options(Config);
            var res = test.GetOptions();
            return res;
        }

        public static List<double> CompteGridDetails(Dictionary<string, Dictionary<string, List<Option>>> res,
            string ticker, out List<double> maturities, out double[,] priceFinalMat)
        {
            var strikes = new List<double>();
            maturities = new List<double>();
            MaturityLines(res, ticker, maturities, strikes);
            strikes.Sort();
            maturities.Sort();
            DataCleanUp(res, ticker, maturities, out priceFinalMat, strikes);
            return strikes;
        }

        private static void MaturityLines(Dictionary<string, Dictionary<string, List<Option>>> res, string ticker,
            List<double> maturities, List<double> strikes)
        {
            foreach (var key in res[ticker].Keys)
                if (res[ticker][key] != null)
                    for (var i = 0; i < res[ticker][key].Count; i++)
                    {
                        var strike = Convert.ToDouble(res[ticker][key][i].strikePrice);
                        var maturity = Convert.ToDouble(res[ticker][key][i].expirationDate);
                        if (!strikes.Contains(strike)) strikes.Add(strike);

                        if (!maturities.Contains(maturity)) maturities.Add(maturity);
                    }
        }

        private static void DataCleanUp(Dictionary<string, Dictionary<string, List<Option>>> res, string ticker,
            List<double> maturities, out double[,] priceFinalMat,
            List<double> strikes)
        {
            var ColumnToRemoveIndex = new List<double>();
            var ColumnToRemoveDate = new List<double>();
            var RowToRemoveIndex = new List<double>();
            var RowToRemoveDate = new List<double>();

            var price = StepOneRemoving(res, ticker, maturities, strikes, ColumnToRemoveIndex,
                ColumnToRemoveDate);
            var priceFinal = new double[strikes.Count, maturities.Count - ColumnToRemoveIndex.Count];

            DeleteColumns(maturities, priceFinal, price, ColumnToRemoveIndex, ColumnToRemoveIndex);

            StepTwoRemoving(priceFinal, strikes, RowToRemoveIndex, RowToRemoveDate);

            priceFinal = ModifyDataInRow(maturities, priceFinal, RowToRemoveIndex, strikes, RowToRemoveDate);

            ModifyDataInColumns(maturities, out priceFinalMat, priceFinal, ColumnToRemoveIndex, ColumnToRemoveDate,
                strikes);
        }

        private static double[,] StepOneRemoving(Dictionary<string, Dictionary<string, List<Option>>> res,
            string ticker, List<double> maturities, List<double> strikes,
            List<double> ColumnToRemoveIndex, List<double> ColumnToRemoveDate)
        {
            int cpt;
            var actualColumn = 0;
            var price = new double[strikes.Count, maturities.Count];
            foreach (var key in res[ticker].Keys.Where(key => res[ticker][key] != null))
            {
                cpt = 0;
                for (var i = 0; i < res[ticker][key].Count; i++)
                {
                    actualColumn = i;
                    var strike = Convert.ToDouble(res[ticker][key][i].strikePrice);
                    var maturity = Convert.ToDouble(res[ticker][key][i].expirationDate);
                    var indexStrike = strikes.IndexOf(strike);
                    var indexMaturity = maturities.IndexOf(maturity);

                    price[indexStrike, indexMaturity] = Convert.ToDouble(res[ticker][key][i].closingPrice);
                    if (price[indexStrike, indexMaturity] > 0) cpt += 1;
                }

                if (cpt < 4)
                {
                    var mat = Convert.ToDouble(res[ticker][key][actualColumn].expirationDate);
                    var indexMaturity = maturities.IndexOf(mat);
                    ColumnToRemoveIndex.Add(indexMaturity);
                    ColumnToRemoveDate.Add(mat);
                }
            }

            return price;
        }

        private static void StepTwoRemoving(double[,] priceFinal, List<double> strikes, List<double> RowToRemoveIndex,
            List<double> RowToRemoveDate)
        {
            var cptRow = 0;
            for (var i = 0; i < priceFinal.GetLength(0); i++)
            {
                cptRow = 0;
                for (var j = 0; j < priceFinal.GetLength(1); j++)
                {
                    var test = priceFinal[i, j];
                    if (test > 0) cptRow += 1;
                }

                if (cptRow < 4)
                    foreach (var k in strikes)
                    {
                        var strike = k;
                        var indexStrike = strikes.IndexOf(strike);
                        if (i == indexStrike)
                        {
                            RowToRemoveIndex.Add(indexStrike);
                            RowToRemoveDate.Add(strike);
                        }
                    }
            }
        }

        private static void ModifyDataInColumns(List<double> maturities, out double[,] priceFinalMat,
            double[,] priceFinal,
            List<double> ColumnToRemoveIndex, List<double> ColumnToRemoveDate, List<double> strikes)
        {
            int cptRow;
            for (var i = 0; i < priceFinal.GetLength(1); i++)
            {
                cptRow = 0;
                for (var j = 0; j < priceFinal.GetLength(0); j++)
                {
                    var test = priceFinal[j, i];
                    if (test > 0) cptRow += 1;
                }

                if (cptRow >= 4) continue;

                foreach (var k in maturities)
                {
                    var mat = k;
                    var indexMat = maturities.IndexOf(mat);
                    if (i == indexMat)
                    {
                        ColumnToRemoveIndex.Add(indexMat);
                        ColumnToRemoveDate.Add(mat);
                    }
                }
            }

            var listMat = new List<List<double>>();
            for (var i = 1; i < priceFinal.GetLength(0) + 1; i++)
            {
                var list1 = new List<double>();
                for (var j = 1; j < priceFinal.GetLength(1) + 1; j++) list1.Add(priceFinal[i - 1, j - 1]);

                listMat.Add(list1);
            }

            foreach (var indice in ColumnToRemoveIndex.OrderByDescending(v => v).Select(item => Convert.ToInt32(item)))
            foreach (var item in listMat)
                item.RemoveAt(indice);


            priceFinalMat = new double[strikes.Count, maturities.Count - ColumnToRemoveIndex.Count];

            foreach (var item in listMat)
            foreach (var item1 in item)
                priceFinalMat[listMat.IndexOf(item), item.IndexOf(item1)] = item1;

            foreach (var item in ColumnToRemoveDate) maturities.Remove(item);
        }

        private static double[,] ModifyDataInRow(List<double> maturities, double[,] priceFinal,
            List<double> RowToRemoveIndex, List<double> strikes,
            List<double> RowToRemoveDate)
        {
            var list = new List<List<double>>();
            for (var i = 1; i < priceFinal.GetLength(0) + 1; i++)
            {
                var list1 = new List<double>();
                for (var j = 1; j < priceFinal.GetLength(1) + 1; j++) list1.Add(priceFinal[i - 1, j - 1]);

                list.Add(list1);
            }

            foreach (var indice in RowToRemoveIndex.OrderByDescending(v => v).Select(item => Convert.ToInt32(item)))
                list.RemoveAt(indice);

            priceFinal = new double[strikes.Count - RowToRemoveIndex.Count, maturities.Count];
            foreach (var item in list)
            foreach (var item1 in item)
                priceFinal[list.IndexOf(item), item.IndexOf(item1)] = item1;

            foreach (var item in RowToRemoveDate) strikes.Remove(item);

            return priceFinal;
        }

        private static void DeleteColumns(List<double> maturities, double[,] priceFinal, double[,] price,
            List<double> RowToRemoveIndex,
            List<double> RowToRemoveDate)
        {
            for (var i = 1; i < price.GetLength(0); i++)
            for (var j = 1; j < price.GetLength(1); j++)
            {
                foreach (var element in RowToRemoveIndex.Where(element => j == element))
                foreach (var date in RowToRemoveDate.Where(date => j == maturities.IndexOf(date)))
                    maturities.Remove(date);
                priceFinal[i - 1, j - 1] = price[i, j];
            }
        }

        public static List<double> MaturitiesFormat(List<double> maturities)
        {
            var newMaturities = new List<double> { };
            string today = DateTime.Today.ToString("dd-MM-yyyy");

            foreach (var mat in maturities)
            {
                var day = GetNumberDay(mat);
                newMaturities.Add(day);
            }
            return newMaturities;
        }

        private static double GetNumberDay(double mat)
        {
            string today = DateTime.Today.ToString("dd-MM-yyyy");
            var result = DateTime.ParseExact(mat.ToString(), "yyyyMMdd", CultureInfo.InvariantCulture).ToString("dd-MM-yyyy");
            TimeSpan diff = Convert.ToDateTime(result) - Convert.ToDateTime(today);
            double day = (diff.TotalDays) / 365;
            day = Math.Round(day, 2);
            return day;
        }
    }
}