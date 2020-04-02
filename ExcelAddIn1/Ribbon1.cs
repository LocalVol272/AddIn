using System;
using System.Collections.Generic;
using Microsoft.Office.Tools.Ribbon;
using ProjetVolSto.PricerObjects;
using ProjetVolSto.Struct;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace ExcelAddIn1
{
    public partial class Ribbon1
    {
        private Worksheet _newWorksheet;

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            List<string> res_string = new List<string>(){"Call","Put"};

            foreach (string value in res_string)
            {
                RibbonDropDownItem item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                item.Label = value;
                comboBox3.Items.Add(item);
            }

            List<string> tickers = new List<string>() { "AAPL", "AMZN", "FB", "GOOG" };

            foreach (string value in tickers)
            {
                RibbonDropDownItem item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
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
        }



        private void FillTicker(string country)
        {
            List<string> res_string = new List<string>();
            switch (country)
            {
                case "us":
                    Stock action = new Stock();
                    action.Token = new Token("Tsk_bbe66f58b6d149f59a9af4eb83bfc7f5");
                    //List<Ticker> res = action.GetAllTickers(country);
                    //res_string = res.ToListString();
                    break;
            }

            comboBox2.Items.Clear();
            foreach (string value in res_string)
            {
                RibbonDropDownItem item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
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

        private void button4_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            string ticker = comboBox2.Text;
            //DL data
            int[] strike_ex = new int[11];
            double[] tenor_ex = new double[11];
            double[,] price_ex = new double[strike_ex.Length,tenor_ex.Length];
            int s = 150;
            double t = 0.25;
            for(int i = 0; i< strike_ex.Length; i++)
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

            _newWorksheet.Range["B6"].Value = "Option Market Price";
            _newWorksheet.Range["B6"].Font.FontStyle = "Bold";
            _newWorksheet.Range["B6"].Font.Underline = true;

            GridView _gv = new GridView(_newWorksheet, strike_ex, tenor_ex);
            _gv.DisplayGrid(7, 3, price_ex);

            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }
    }
}
