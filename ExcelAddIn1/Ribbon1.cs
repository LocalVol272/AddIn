using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using ExcelAddIn1.PricerObjects;
using ExcelAddIn1.PricingCalculation;
using Extreme.Mathematics.Calculus.OrdinaryDifferentialEquations;
using MathNet.Numerics;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using Newtonsoft.Json.Linq;

namespace ExcelAddIn1
{
    public partial class Ribbon1
    {
        private int _lastRow;
        private Worksheet _newWorksheet;
        private double _spot;
        private string _type;
        private double _r;
        private double _moneyness = 2.5;
        private double [] _strikes;
        private double[] _mStrikes;
        private double [] _tenors;
        private double [,] _optionMarketPrice;
        private double[,] _mOptionMarketPrice;
        private double[,] _VolLocale;

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            var resString = new List<string> {"Call", "Put"};

            foreach (var value in resString)
            {
                var item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                item.Label = value;
                comboBox3.Items.Add(item);
            }

            var tickers = new List<string> {"AAPL", "AMZN", "FB", "GOOG"};

            foreach (var value in tickers)
            {
                var item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                item.Label = value;
                comboBox2.Items.Add(item);
            }
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            CleanRibbon();
            _newWorksheet = Globals.ThisAddIn.Application.Worksheets.Add();
            NewWorksheet.Creation(_newWorksheet, DateTime.Now.ToString("HH:mm:ss"));

            _newWorksheet.EnableCalculation = true;
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }

        private void CleanRibbon()
        {
            comboBox3.Text = "";
            comboBox2.Text = "";
        }

        private void editBox1_TextChanged(object sender, RibbonControlEventArgs e)
        {
        }

        private void Price_Click(object sender, RibbonControlEventArgs e)
        {
            
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            _newWorksheet.Range["B" + _lastRow].Value = "Volatility Surface";
            _newWorksheet.Range["B" + _lastRow].Font.FontStyle = "Bold";
            _newWorksheet.Range["B" + _lastRow].Font.Underline = true;
            applyMoneyness();
            Grid grid = new Grid(_mOptionMarketPrice, _tenors, _mStrikes);
            _VolLocale = grid.LocalVolatility(_mOptionMarketPrice, _mStrikes, _tenors, _r);
            var gv = new GridView(_newWorksheet, _mStrikes, _tenors);
            gv.DisplayGrid(_lastRow + 1, 3, _VolLocale);
            gv.DisplayVolSurface("Volatility Surface", _lastRow + 2, 4);
            _lastRow += _mStrikes.Length + 2;
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }

        private void applyMoneyness()
        {
            double borneInf;
            int indexBorneInf;
            double borneSup;
            int indexBorneSup;
            borneInf = _spot * (1 - _moneyness);
            if(borneInf < 0)
            {
                borneInf = 0;
            }
            borneSup = _spot * (1 + _moneyness);
            int i = 0;
            while(_strikes[i] <= borneInf)
            {
                i++;
            }
            borneInf = _strikes[i];
            indexBorneInf = i;
            int j = _strikes.Length-1;
            while (_strikes[j] >= borneSup)
            {
                j--;
            }
            borneSup = _strikes[j];
            indexBorneSup = j;
            _mStrikes = new double[j - i + 1];
            _mOptionMarketPrice = new double[j - i + 1, _tenors.Length];
            for(int k = indexBorneInf; k <= indexBorneSup; k++)
            {
                _mStrikes[k - indexBorneInf] = _strikes[k];
                for(int t = 0; t< _tenors.Length; t++)
                {
                    _mOptionMarketPrice[k - indexBorneInf, t] = _optionMarketPrice[k, t];
                }
            }
        }

        private void FillTicker(string country)
        {
            var resString = new List<string>();
            switch (country)
            {
                case "us":
                    var action = new Stock();
                    action.Token = new Token("Tsk_bbe66f58b6d149f59a9af4eb83bfc7f5");
                    //List<Ticker> res = action.GetAllTickers(country);
                    //res_string = res.ToListString();
                    break;
            }

            comboBox2.Items.Clear();
            foreach (var value in resString)
            {
                var item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                item.Label = value;
                comboBox2.Items.Add(item);
            }
        }

        private void comboBox2_TextChanged(object sender, RibbonControlEventArgs e)
        {
            _newWorksheet.Range["B1"].Value = comboBox2.Text;
        }

        private void comboBox3_TextChanged(object sender, RibbonControlEventArgs e)
        {
            _type = comboBox3.Text;
            _newWorksheet.Range["B3"].Value = _type;
        }

        private void button4_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            var ticker = comboBox2.Text;
            var action = new Stock(ticker);
            action.Token = new Token("Tsk_bbe66f58b6d149f59a9af4eb83bfc7f5");
            try
            {
                if (_newWorksheet == null) throw new Exception("WORKSHEET");
                if (_newWorksheet.Range["B1"].Value == null) throw new Exception("TICKER");
                if (_newWorksheet.Range["B3"].Value == null) throw new Exception("OPTION");
                _spot = action.GetLastPrice();
            }
            catch (Exception exception)
            {
                switch (exception.Message)
                {
                    case "WORKSHEET":
                        MessageBox.Show("Merci de créer une nouvelle feuille.");
                        Globals.ThisAddIn.Application.ScreenUpdating = true;
                        return;
                    case "TICKER":
                        MessageBox.Show("Vous devez saisir un ticker.");
                        Globals.ThisAddIn.Application.ScreenUpdating = true;
                        return;
                    case "OPTION":
                        MessageBox.Show("Vous devez saisir le type d'Option(Call ou Put).");
                        Globals.ThisAddIn.Application.ScreenUpdating = true;
                        return;
                    default:
                        MessageBox.Show("Il y a une erreur");
                        return;
                }
            }

            _newWorksheet.Range["B2"].Value = _spot;

            var res = GetOptions(ticker);

            var strikes = CompteGridDetails(res, ticker, out var maturities, out _optionMarketPrice);
            _strikes = strikes.Select(item => Convert.ToDouble(item)).ToArray();
            _tenors = MaturitiesFormat(maturities).Select(item => Convert.ToDouble(item)).ToArray();

            _newWorksheet.Range["B6"].Value = "Option Market Price";
            _newWorksheet.Range["B6"].Font.FontStyle = "Bold";
            _newWorksheet.Range["B6"].Font.Underline = true;

            var gv = new GridView(_newWorksheet, _strikes, _tenors);
            gv.DisplayGrid(7, 3, _optionMarketPrice);
     
            _lastRow = 11 + _strikes.Length;
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }

        private static List<double> MaturitiesFormat(List<double> maturities)
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

        private static List<double> CompteGridDetails(Dictionary<string, Dictionary<string, List<Option>>> res, string ticker, out List<double> maturities, out double[,] priceFinal)
        {
            List<double> strikes = new List<double>();
            maturities = new List<double>();

            foreach (string key in res[ticker].Keys)
            {
                if (res[ticker][key] != null)
                {
                    for (int i = 0; i < res[ticker][key].Count; i++)
                    {
                        double strike = Convert.ToDouble(res[ticker][key][i].strikePrice);
                        double maturity = Convert.ToDouble(res[ticker][key][i].expirationDate);
                        if (!strikes.Contains(strike))
                        {
                            strikes.Add(strike);
                        }

                        if (!maturities.Contains(maturity))
                        {
                            maturities.Add(maturity);
                        }
                    }
                }
            }

            strikes.Sort();
            maturities.Sort();
            var cpt = 0;
            var actualColumn = 0;
            List<double> ColumnToRemoveIndex = new List<double>();
            List<double> ColumnToRemoveDate = new List<double>();
            List<double> RowToRemoveIndex = new List<double>();
            List<double> RowToRemoveDate = new List<double>();
            var price = new double[strikes.Count, maturities.Count];
            foreach (var key in res[ticker].Keys.Where(key => res[ticker][key] != null))
            {
                cpt = 0;
                for (int i = 0; i < res[ticker][key].Count; i++)
                {
                    actualColumn = i;
                    double strike = Convert.ToDouble(res[ticker][key][i].strikePrice);
                    double maturity = Convert.ToDouble(res[ticker][key][i].expirationDate);
                    int indexStrike = strikes.IndexOf(strike);
                    int indexMaturity = maturities.IndexOf(maturity);
   
                    price[indexStrike, indexMaturity] = Convert.ToDouble(res[ticker][key][i].closingPrice);
                    if (price[indexStrike, indexMaturity] > 0)
                    {
                        cpt += 1;
                    }
                }

                if (cpt < 4)
                {
                    double mat = Convert.ToDouble(res[ticker][key][actualColumn].expirationDate);
                    int indexMaturity = maturities.IndexOf(mat);
                    ColumnToRemoveIndex.Add(indexMaturity);
                    ColumnToRemoveDate.Add(mat);
                }
            }


            priceFinal = new double[strikes.Count, maturities.Count- ColumnToRemoveIndex.Count];

            DeleteColumns(maturities, priceFinal, price, ColumnToRemoveIndex, ColumnToRemoveIndex);
            int cptRow = 0;
            for (int i = 0; i < priceFinal.GetLength(0); i++)
            {
                cptRow = 0;
                for (int j = 0; j < priceFinal.GetLength(1); j++)
                {
                    double test = priceFinal[i, j];
                    if (test > 0)
                    {
                        cptRow += 1;
                    }
                }
                if (cptRow < 4)
                {
                    foreach (var k in strikes)
                    {
                        double strike = k;
                        int indexStrike = strikes.IndexOf(strike);
                        if (i == indexStrike)
                        {
                            RowToRemoveIndex.Add(indexStrike);
                            RowToRemoveDate.Add(strike);
                        }
                    }
                }
            }
            

            List<List<double>> list = new List<List<double>>();
            for (int i = 1; i < priceFinal.GetLength(0)+1; i++)
            {
                List<double> list1 = new List<double>();
                for (int j = 1; j < priceFinal.GetLength(1)+1; j++)
                {
                    list1.Add(priceFinal[i-1, j-1]);
                }
                list.Add(list1);
            }

            foreach (var indice in RowToRemoveIndex.OrderByDescending(v => v).Select((item => Convert.ToInt32(item))))
            {
                list.RemoveAt(indice);
            }

            priceFinal = new double[strikes.Count - RowToRemoveIndex.Count, maturities.Count];
            foreach (var item in list)
            {
                foreach (var item1 in item)
                {
                    priceFinal[list.IndexOf(item), item.IndexOf(item1)] = item1;
                }
            }

            foreach (var item in RowToRemoveDate)
            {
                    strikes.Remove(item);
            }

            //var priceFinal2 = new double[strikes.Count - RowToRemoveIndex.Count, maturities.Count];
            //DeleteRows(strikes, priceFinal2, priceFinal, RowToRemoveIndex, RowToRemoveDate);

            return strikes;
        }

        private static void DeleteRows(List<double> strikes, double[,] priceFinal, double[,] price, List<double> RowToRemoveIndex,
            List<double> RowToRemoveDate)
        {
            int u=1;
            int cpt = -1;
            for (int i = 1; i < price.GetLength(0); i++)
            {
                bool existR = false;
                foreach (var element in RowToRemoveIndex.Where(element => (i-1) == element))
                {
                    foreach (var date in RowToRemoveDate.Where(date => (i - 1) == strikes.IndexOf(date)))
                    {
                        strikes.Remove(date);
                        existR = true;
                        u = u + 1;
                        break;
                    }
                }

                if (existR.Equals(true))
                    continue;

                cpt = cpt + 1;
                for (int j = 1; j < price.GetLength(1); j++)
                {
                    
                    priceFinal[cpt, j - 1] = price[i, j];
                }
            }
        }

        private static void DeleteColumns(List<double> maturities, double[,] priceFinal, double[,] price, List<double> RowToRemoveIndex,
            List<double> RowToRemoveDate)
        {
            for (int i = 1; i < price.GetLength(0); i++)
            {
                for (int j = 1; j < price.GetLength(1); j++)
                {
                    foreach (var element in RowToRemoveIndex)
                    {
                        if (j == element)
                        {
                            foreach (var date in RowToRemoveDate)
                            {
                                if (j == maturities.IndexOf(date))
                                {
                                    maturities.Remove(date);
                                }
                            }

                            continue;
                        }
                    }

                    priceFinal[i - 1, j - 1] = price[i, j];
                }
            }
        }

        private Dictionary<string, Dictionary<string, List<Option>>> GetOptions(string ticker)
        {
            List<string> TickerList = new List<string>() {ticker};
            Dictionary<string, object> Params = new Dictionary<string, object>();
            List<string> DateList = new List<string>() { };
            
            Params.Add("ProductType", "Option/" + comboBox3.Text);
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

        private static double GetNumberDay(double mat)
        {
            string today = DateTime.Today.ToString("dd-MM-yyyy");
            var result = DateTime.ParseExact(mat.ToString(), "yyyyMMdd", CultureInfo.InvariantCulture).ToString("dd-MM-yyyy");
            TimeSpan diff = Convert.ToDateTime(result) - Convert.ToDateTime(today);
            double day = (diff.TotalDays) / 365;
            day = Math.Round(day, 2);
            return day;
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            Grid grid = new Grid(_VolLocale, _tenors, _mStrikes);
            double[,] prices = grid.BSPD(_spot, _r, _VolLocale, _type);

            _lastRow += 3;
            _newWorksheet.Range["B" + _lastRow].Value = "Option Price with Local Volatility";
            _newWorksheet.Range["B" + _lastRow].Font.FontStyle = "Bold";
            _newWorksheet.Range["B" + _lastRow].Font.Underline = true;

            var gv = new GridView(_newWorksheet, _mStrikes, _tenors);
            gv.DisplayGrid(_lastRow + 1, 3, prices);


            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }

        private void editBox1_TextChanged_1(object sender, RibbonControlEventArgs e)
        {
            _moneyness = Convert.ToDouble(editBox1.Text);
        }

        private void editBox2_TextChanged(object sender, RibbonControlEventArgs e)
        {
            _r = Convert.ToDouble(editBox2.Text);
        }

        private void button5_Click(object sender, RibbonControlEventArgs e)
        {

        }
    }
}