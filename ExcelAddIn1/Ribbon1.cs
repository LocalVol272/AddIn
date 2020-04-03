using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ExcelAddIn1.PricerObjects;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;

namespace ExcelAddIn1
{
    public partial class Ribbon1
    {
        private int _lastRow;
        private Worksheet _newWorksheet;
        private double _spot;
        private string _type;
        private double _r;
        private double _moneyness;


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
            _newWorksheet = Globals.ThisAddIn.Application.Worksheets.Add();
            NewWorksheet.Creation(_newWorksheet, DateTime.Now.ToString("HH:mm:ss"));

            _newWorksheet.EnableCalculation = true;
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }

        private void editBox1_TextChanged(object sender, RibbonControlEventArgs e)
        {
        }

        private void Price_Click(object sender, RibbonControlEventArgs e)
        {
            //ICI ON GO POUR PRICER


            /*            List<string> TickerList = new List<string>() { "AAPL","FB","TSLA" };
                        Dictionary<string, object> Params = new Dictionary<string, object>();
                        List<string> DateList = new List<string>(){"20201030","202010","202002"};
                        Params.Add("ProductType", "Option/Call");
                        Params.Add("Dates", DateList);
                        Params.Add("Tickers", TickerList);

                        Dictionary<string, object> Config = new Dictionary<string, object>(){};
                        Config.Add("Token","Tsk_bbe66f58b6d149f59a9af4eb83bfc7f5");
                        Config.Add("Type", "GET");
                        Config.Add("Params", Params);

                        Options test = new Options(Config);
                        Dictionary<string, Dictionary<string, List<Option>>> res = test.GetOptions();
                        Config.Add("TYPE", "GET");*/
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            _newWorksheet.Range["B" + _lastRow].Value = "Volatility Surface";
            _newWorksheet.Range["B" + _lastRow].Font.FontStyle = "Bold";
            _newWorksheet.Range["B" + _lastRow].Font.Underline = true;
            //DL data
            var strike_ex = new double[11];
            var tenor_ex = new double[11];
            var price_ex = new double[strike_ex.Length, tenor_ex.Length];
            var s = 150;
            var t = 0.25;
            for (var i = 0; i < strike_ex.Length; i++)
            {
                strike_ex[i] = s;
                tenor_ex[i] = t;
                s += 10;
                t += 0.25;
                var c = 10;
                for (var j = 0; j < tenor_ex.Length; j++)
                {
                    c += 2;
                    price_ex[i, j] = s * t + c;
                }
            }

            var gv = new GridView(_newWorksheet, strike_ex, tenor_ex);
            gv.DisplayGrid(_lastRow + 1, 3, price_ex);
            gv.DisplayVolSurface("Volatility Surface", _lastRow + 2, 4);
            _lastRow += strike_ex.Length + 2;
            Globals.ThisAddIn.Application.ScreenUpdating = true;
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
            _spot = action.GetLastPrice();
            _newWorksheet.Range["B2"].Value = _spot;

            var res = GetOptions(ticker);

            var strikes = CompteGridDetails(res, ticker, out var maturities, out var priceData);
            var strikeData = strikes.Select(item => Convert.ToDouble(item)).ToArray();
            var tenorData = MaturitiesFormat(maturities).Select(item => Convert.ToDouble(item)).ToArray();

            _newWorksheet.Range["B6"].Value = "Option Market Price";
            _newWorksheet.Range["B6"].Font.FontStyle = "Bold";
            _newWorksheet.Range["B6"].Font.Underline = true;

            var gv = new GridView(_newWorksheet, strikeData, tenorData);
            gv.DisplayGrid(7, 3, priceData);
     
            _lastRow = 11 + strikeData.Length;
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

        private static List<double> CompteGridDetails(Dictionary<string, Dictionary<string, List<Option>>> res, string ticker, out List<double> maturities, out double[,] price)
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

            price = new double[strikes.Count, maturities.Count];
            foreach (string key in res[ticker].Keys)
            {
                if (res[ticker][key] != null)
                {
                    for (int i = 0; i < res[ticker][key].Count; i++)
                    {
                        double strike = Convert.ToDouble(res[ticker][key][i].strikePrice);
                        double maturity = Convert.ToDouble(res[ticker][key][i].expirationDate);
                        int indexStrike = strikes.IndexOf(strike);
                        int indexMaturity = maturities.IndexOf(maturity);
                        price[indexStrike, indexMaturity] = Convert.ToDouble(res[ticker][key][i].closingPrice);
                    }
                }
            }

            return strikes;
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
            //DL data
            var strike_ex = new double[11];
            var tenor_ex = new double[11];
            var price_ex = new double[strike_ex.Length, tenor_ex.Length];
            var s = 150;
            var t = 0.25;
            for (var i = 0; i < strike_ex.Length; i++)
            {
                strike_ex[i] = s;
                tenor_ex[i] = t;
                s += 10;
                t += 0.25;
                var c = 10;
                for (var j = 0; j < tenor_ex.Length; j++)
                {
                    c += 2;
                    price_ex[i, j] = s * t + c;
                }
            }

            _lastRow += 3;
            _newWorksheet.Range["B" + _lastRow].Value = "Option Price with Local Volatility";
            _newWorksheet.Range["B" + _lastRow].Font.FontStyle = "Bold";
            _newWorksheet.Range["B" + _lastRow].Font.Underline = true;

            var gv = new GridView(_newWorksheet, strike_ex, tenor_ex);
            gv.DisplayGrid(_lastRow + 1, 3, price_ex);


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
    }
}