using System;
using System.Collections.Generic;
using ExcelAddIn1.PricerObjects;
using Microsoft.Office.Tools.Ribbon;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace ExcelAddIn1
{
    public partial class Ribbon1
    {
        private Worksheet _newWorksheet;

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            var resString = new List<string>() {"Call", "Put"};

            foreach (var value in resString)
            {
                var item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                item.Label = value;
                comboBox3.Items.Add(item);
            }

            var tickers = new List<string>() {"AAPL", "AMZN", "FB", "GOOG"};

            foreach (var value in tickers)
            {
                var item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                item.Label = value;
                comboBox2.Items.Add(item);
            }
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            _newWorksheet = Globals.ThisAddIn.Application.Worksheets.Add();
            NewWorksheet.Creation(_newWorksheet, DateTime.Now.ToString("HH:mm:ss"));

            _newWorksheet.EnableCalculation = true;
        }

        private void editBox1_TextChanged(object sender, RibbonControlEventArgs e)
        {
        }

        private void Price_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            _newWorksheet.Range["B" +_lastRow].Value = "Volatility Surface";
            _newWorksheet.Range["B" + _lastRow].Font.FontStyle = "Bold";
            _newWorksheet.Range["B" + _lastRow].Font.Underline = true;
            //DL data
            int[] strike_ex = new int[11];
            double[] tenor_ex = new double[11];
            double[,] price_ex = new double[strike_ex.Length, tenor_ex.Length];
            int s = 150;
            double t = 0.25;
            for (int i = 0; i < strike_ex.Length; i++)
            {
                strike_ex[i] = s;
                tenor_ex[i] = t;
                s += 10;
                t += 0.25;
                int c = 10;
                for (int j = 0; j < tenor_ex.Length; j++)
                {
                    c += 2;
                    price_ex[i, j] = s * t + c;
                }
            }
            GridView gv = new GridView(_newWorksheet, strike_ex, tenor_ex);
            gv.DisplayGrid(_lastRow + 1, 3, price_ex);
            gv.DisplayVolSurface("Volatility Surface", _lastRow + 2, 4);
            _lastRow +=  strike_ex.Length + 2;
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
            _newWorksheet.Range["B3"].Value = comboBox3.Text;
        }
        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            //DL data
            int[] strike_ex = new int[11];
            double[] tenor_ex = new double[11];
            double[,] price_ex = new double[strike_ex.Length, tenor_ex.Length];
            int s = 150;
            double t = 0.25;
            for (int i = 0; i < strike_ex.Length; i++)
            {
                strike_ex[i] = s;
                tenor_ex[i] = t;
                s += 10;
                t += 0.25;
                int c = 10;
                for (int j = 0; j < tenor_ex.Length; j++)
                {
                    c += 2;
                    price_ex[i, j] = s * t + c;
                }
            }
            _lastRow += 3;
            _newWorksheet.Range["B"+_lastRow].Value = "Option Price with Local Volatility";
            _newWorksheet.Range["B" + _lastRow ].Font.FontStyle = "Bold";
            _newWorksheet.Range["B" + _lastRow ].Font.Underline = true;

            GridView gv = new GridView(_newWorksheet, strike_ex, tenor_ex);
            gv.DisplayGrid(_lastRow+1, 3, price_ex);
           

            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }
        private void button4_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            var ticker = comboBox2.Text;
            //DL data
            var strike_ex = new int[11];
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

            _newWorksheet.Range["B6"].Value = "Option Market Price";
            _newWorksheet.Range["B6"].Font.FontStyle = "Bold";
            _newWorksheet.Range["B6"].Font.Underline = true;

            var gv = new GridView(_newWorksheet, strike_ex, tenor_ex);
            gv.DisplayGrid(7, 3, price_ex);

            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }

    }
}
