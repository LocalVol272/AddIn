using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace ExcelAddIn1.PricerObjects
{
    public static class TickerFormat
    {
        public static List<string> ToListString(this List<Ticker> list)
        {
            var listTickers = new List<string>();
            list.ForEach(x => listTickers.Add(x.symbol));
            return listTickers;
        }
    }

    internal class Stock : DataLoader, IAuthentification
    {
        private YahooRequest _requestContent;
        private readonly string _ticker;
        private new HttpsRequest request;
        private string url;

        public Stock(Dictionary<string, object> config)
        {
            Config = config;
            Token = GetToken(config);
            _request = new ApiRequest();
        }

        public Stock()
        {
            _request = new ApiRequest();
        }

        public Stock(string ticker)
        {
            _ticker = ticker;
        }

        private string Reponse { get; set; }

        public YahooRequest RequestContent
        {
            get => _requestContent;
            set => RequestContent = _requestContent;
        }

        public Token Token { get; set; }

        public Token GetToken(Dictionary<string, object> config)
        {
            var Token = "Token";

            if (config.ContainsKey(Token)) return new Token(config[Token].ToString());
            throw new Exception(string.Format(ConfigError.MissingKey, Token));
        }


        public bool Authentification(Token token)
        {
            if (token.value is null)
                throw new Exception(ConfigError.MissingTokenValue);
            return true;
        }


        public Dictionary<string, Dictionary<string, List<Dictionary<string, object>>>> GetAllTickers(string country)
        {
            string[] args = {country};
            var stack = new StackTrace();
            var root = stack.GetFrame(0).GetMethod().Name;
            Init(args, root);
            GetReponse();
            //FormatOption(JsonConvert.DeserializeObject<Dictionary<string, Dictionary<string, List<Dictionary<string, object>>>>>(_response));
            return JsonConvert
                .DeserializeObject<Dictionary<string, Dictionary<string, List<Dictionary<string, object>>>>>(Reponse);
        }


        public double GetLastPrice()
        {
            var yesterday = DateTime.Now.Date.Subtract(TimeSpan.FromDays(1)).ConvertToTimestamp().ToString();
            var today = DateTime.Now.ConvertToTimestamp().ToString();
            string[] args = {_ticker, yesterday, today};
            var stack = new StackTrace();
            var root = stack.GetFrame(0).GetMethod().Name;
            Init(args, root);
            GetReponse();
            var historicalData = JsonConvert.DeserializeObject<YahooChartObject>(Reponse);
            var HistoPrices = historicalData.chart.result[0].indicators.adjclose[0].adjclose;

            return (double) HistoPrices[HistoPrices.Count - 1];
        }

        public static List<string> GetAllTickers()
        {
            return AvailableData.Ticker;
        }

        private void FormatTickers()
        {
        }

        private void GetReponse()
        {
            Reponse = ExecuteRequest(url)
                .GetAwaiter()
                .GetResult();
            _requestContent.Response = Reponse;
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
            return await request.Post(url, requestContent);
        }


        private void BuildUrl(string root, [Optional] string[] args)
        {
            switch (root)
            {
                case "GetLastPrice":
                    url = string.Format(ApiMapping.Roots[root], args[0], args[1], args[2]);


                    break;
            }
        }


        private void InitRequest()
        {
            if (Authentification(Token))
            {
                switch (Request)
                {
                    case null:
                        Request = new ApiRequest();
                        break;
                }

                request = new HttpsRequest();
                Request.RequestContent = new YahooRequest();
            }
        }
    }
}