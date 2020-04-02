using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace ExcelAddIn1.PricerObjects
{
    public class YahooChartObject
    {
        public YahooChartPrices chart { get; set; }
    }


    public class YahooChartPrices
    {
        public List<YahooChartInfo> result { get; set; }
    }


    public class YahooChartInfo
    {
        public List<double> timestamp { get; set; }
        public Indicators indicators { get; set; }
    }

    public class Indicators
    {
        public List<AdjClosePrices> adjclose { get; set; }
    }

    //public class AdjClose
    //{
    //    public List<AdjClosePrices> adjclose { get; set; }
    //}

    public class AdjClosePrices
    {
        public List<object> adjclose { get; set; }
    }


    public class Ticker
    {
        public string symbol { get; set;}
    }

    public static class TickerFormat
    {
        public static List<string> ToListString(this List<Ticker> list)
        {
            List<string> listTickers = new List<string>();
            list.ForEach(x => listTickers.Add(x.symbol));
            return listTickers;
        }

    }




    class Stock : DataLoader, IAuthentification
    {
        private string _response;
        private Token _token;
        private string url;
        
        private YahooRequest _requestContent;
        private string _ticker;
        private string Reponse { get => _response; set => _response = value; }
        private new HttpsRequest request;
        public YahooRequest RequestContent { get => _requestContent; set => RequestContent = _requestContent; }
        public Token Token { get =>_token; set => _token=value; }
        public Stock(Dictionary<string, object> config)
        {
            this.Config = config;
            this.Token = GetToken(config);
            _request = new ApiRequest();
        }

        public Stock() => _request = new ApiRequest();
        public Stock(string ticker) =>_ticker = ticker;




        public Dictionary<string, Dictionary<string, List<Dictionary<string, object>>>> GetAllTickers(string country)
        {
            string[] args = { country };
            var stack = new StackTrace();
            string root = stack.GetFrame(0).GetMethod().Name;
            Init(args, root);
            GetReponse();
            //FormatOption(JsonConvert.DeserializeObject<Dictionary<string, Dictionary<string, List<Dictionary<string, object>>>>>(_response));
            return  JsonConvert.DeserializeObject<Dictionary<string, Dictionary<string, List<Dictionary<string, object>>>>>(_response);
        }


        public double GetLastPrice()
        {
            string yesterday = DateTime.Now.Date.Subtract(TimeSpan.FromDays(1)).ConvertToTimestamp().ToString();
            string today = DateTime.Now.ConvertToTimestamp().ToString();
            string[] args = { _ticker,yesterday,today};
            var stack = new StackTrace();
            string root = stack.GetFrame(0).GetMethod().Name;
            Init( args,root);
            GetReponse();
            YahooChartObject historicalData = JsonConvert.DeserializeObject<YahooChartObject>(_response);
            List<object> HistoPrices = historicalData.chart.result[0].indicators.adjclose[0].adjclose;
    
            return (double)HistoPrices[HistoPrices.Count - 1];
            
        }

        public static List<string> GetAllTickers()
        {
            return AvailableData.GetTicker();
        }

        private void FormatTickers()
        {
                        
        }

        private void GetReponse()
        {
            _response = ExecuteRequest(url)
                            .GetAwaiter()
                            .GetResult();
            _requestContent.Response = _response;
        }

        private void Init(string[] args, string root)
        {
            InitRequest();
            BuildUrl(root, args);
        }

        private async Task<string> ExecuteRequest(string url)
        {

            return await request.Get(url);
        }


        private async Task<string> ExecuteRequest(string url, HttpContent requestContent)
        {

            return await request.Post(url,requestContent);
        }


        private void BuildUrl(string root,[Optional]string[] args)
        {

            switch (root)
            {

                case "GetLastPrice":
                    url = String.Format(ApiMapping.Roots[root], args[0], args[1],args[2]);



                    break;
                
            }

        }


        private void InitRequest()
        {
            if (Authentification(Token))
            {
                switch(Request)
                {
                    case null:
                        Request = new ApiRequest();
                        break;

                }    
                request = new HttpsRequest();
                Request.RequestContent = new YahooRequest();
            }

        }

        public Token GetToken(Dictionary<string, object> config)
        {
            string Token = "Token";

            if (config.ContainsKey(Token))
            {
                return new Token(config[Token].ToString());
            }
            throw new Exception(String.Format(ConfigError.MissingKey, Token));
        }


        public bool Authentification(Token token)
        {
            if (token.value is null) { throw new Exception(ConfigError.MissingTokenValue); }
            else { return true; }
        }



        
    }


}

