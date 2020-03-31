using System;
using System.Collections.Generic;
using Microsoft.Office.Tools.Ribbon;
using ProjetVolSto.Struct;
using ProjetVolSto.PricerObjects;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace ExcelAddIn1
{
    public partial class Ribbon1
    {
        private string _Country;
        private Worksheet _newWorksheet;

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            List<string> res_string = new List<string>(){"us","fra"};

            foreach (string value in res_string)
            {
                RibbonDropDownItem item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                item.Label = value;
                comboBox1.Items.Add(item);
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

        private void comboBox1_TextChanged(object sender, RibbonControlEventArgs e)
        {
            _Country = comboBox1.Text;
            FillTicker(_Country);
        }

        private void FillTicker(string country)
        {
            List<string> res_string = new List<string>();
            switch (country)
            {
                case "us":
                    Stock action = new Stock();
                    action.Token = new Token("Tsk_bbe66f58b6d149f59a9af4eb83bfc7f5");
                    List<Ticker> res = action.GetAllTickers(country);
                    res_string = res.ToListString();
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
            _newWorksheet.Range["B4"].Value = comboBox2.Text;
        }
    }
}