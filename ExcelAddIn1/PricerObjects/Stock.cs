using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace ExcelAddIn1.PricerObjects
{
    public class StockPrice
    {
    }


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
        private IEXRequest _requestContent;

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

        public Stock(Token token)
        {
            Token = token;
        }

        private string Reponse { get; set; }

        public IEXRequest RequestContent
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


        public static List<string> GetAllTickers()
        {
            return AvailableData.GetTicker();
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
                case "GetAllTickers":
                    url = "https://query1.finance.yahoo.com/v7/finance/options/MSFT";
                    break;
            }
        }

        private void InitRequest()
        {
            if (Authentification(Token))
            {
                request = new HttpsRequest();
                Request.RequestContent = new IEXRequest();
            }
        }
    }
}