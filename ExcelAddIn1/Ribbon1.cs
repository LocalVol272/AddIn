using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using ExcelAddIn1.PricerObjects;
using ExcelAddIn1.PricingCalculation;
using Microsoft.Office.Tools.Ribbon;

namespace ExcelAddIn1
{
    public partial class Ribbon1
    {
        private Parameters _details ;
        private bool _step4done;
        private bool _step5done;

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            var resString = new List<string> {"Call", "Put"};
            _details = new Parameters();

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
            _details = new Parameters();
            _step4done = false;
            CleanRibbon();
            _details.newWorksheet = Globals.ThisAddIn.Application.Worksheets.Add();
            NewWorksheet.Creation(_details.newWorksheet);
            _details.newWorksheet.EnableCalculation = true;
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }

        private void CleanRibbon()
        {
            comboBox3.Text = "";
            comboBox2.Text = "";
        }

        private void Price_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            if (_step4done)
            {
                try
                {
                    VolatilyCalculation.VolatilyMain(_details);
                    _step5done = true;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            else
            {
                MessageBox.Show("Vous devez d'abord récupérer les prix des options en cliquant sur 'Import Data''");
            }

            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }

        private void comboBox2_TextChanged(object sender, RibbonControlEventArgs e)
        {
            _details.newWorksheet.Range["B1"].Value = comboBox2.Text;
        }

        private void comboBox3_TextChanged(object sender, RibbonControlEventArgs e)
        {
            _details.type = comboBox3.Text;
            _details.newWorksheet.Range["B3"].Value = _details.type;
        }

        private void button4_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            var ticker = comboBox2.Text;
            var action = new Stock(ticker);
            action.Token = new Token("Tsk_bbe66f58b6d149f59a9af4eb83bfc7f5");
            if (SecureWorksheet.SecuriseNewWs(_details)) return;
            _details.spot = action.GetLastPrice();
            _details.newWorksheet.Range["B2"].Value = _details.spot;

            var res = PriceView.GetOptions(ticker,comboBox3.Text);
            var strikes = PriceView.CompteGridDetails(res, ticker, out var maturities, out _details.optionMarketPrice);
            _details.strikes = strikes.Select(item => Convert.ToDouble(item)).ToArray();
            _details.tenors = PriceView.MaturitiesFormat(maturities).Select(item => Convert.ToDouble(item)).ToArray();

            _details.newWorksheet.Range["B6"].Value = "Option Market Price";
            _details.newWorksheet.Range["B6"].Font.FontStyle = "Bold";
            _details.newWorksheet.Range["B6"].Font.Underline = true;

            var gv = new GridView(_details.newWorksheet, _details.strikes, _details.tenors);
            gv.DisplayGrid(7, 3, _details.optionMarketPrice);
     
            _details.lastRow = 11 + _details.strikes.Length;
            _step4done = true;
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            if (_step5done)
            {

                Globals.ThisAddIn.Application.ScreenUpdating = false;
            Grid grid = new Grid(_details.VolLocale, _details.tenors, _details.mStrikes);
            double[,] prices = grid.BSPD(_details.spot, _details.r, _details.VolLocale, _details.type);

            _details.lastRow += 3;
            _details.newWorksheet.Range["B" + _details.lastRow].Value = "Option Price with Local Volatility";
            _details.newWorksheet.Range["B" + _details.lastRow].Font.FontStyle = "Bold";
            _details.newWorksheet.Range["B" + _details.lastRow].Font.Underline = true;

            var gv = new GridView(_details.newWorksheet, _details.mStrikes, _details.tenors);
            gv.DisplayGrid(_details.lastRow + 1, 3, prices);

            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }
        else
        {
            MessageBox.Show("Veuillez calculer la volatlité locale avant de pricer.");
        }
        }

        private void editBox1_TextChanged_1(object sender, RibbonControlEventArgs e)
        {
            var value = editBox1.Text.Replace('.', ',');
            try
            {
                _details.moneyness = Convert.ToDouble(value);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Assurez vous de rentrer un nombre.");
            }
        }

        private void editBox2_TextChanged(object sender, RibbonControlEventArgs e)
        {
            var value = editBox1.Text.Replace('.', ',');
            try
            {
                _details.r = Convert.ToDouble(value);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Assurez vous de rentrer un nombre.");
            }
        }
    }
}