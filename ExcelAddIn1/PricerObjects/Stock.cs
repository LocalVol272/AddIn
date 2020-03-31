using ProjetVolSto.Struct;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Net.Http;
using Newtonsoft.Json;
using System.Runtime.InteropServices;

namespace ProjetVolSto.PricerObjects
{
    public class Ticker
    {
        public string symbol { get; set; }
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
        private string url;
        private Token _token;
        private IEXRequest _requestContent;

        public Token Token
        {
            get => _token;
            set => _token = value;
        }

        private string Reponse
        {
            get => _response;
            set => _response = value;
        }

        private new HttpsRequest request;

        public IEXRequest RequestContent
        {
            get => _requestContent;
            set => RequestContent = _requestContent;
        }

        public Stock(Dictionary<string, object> config)
        {
            this.Config = config;
            this.Token = GetToken(config);
            _request = new ApiRequest();
        }

        public Stock() => _request = new ApiRequest();
        public Stock(Token token) => Token = token;

        public List<Ticker> GetAllTickers(string country)
        {
            string[] args = {country};
            var stack = new StackTrace();
            string root = stack.GetFrame(0).GetMethod().Name;
            Init(args, root);
            GetReponse();

            return JsonConvert.DeserializeObject<List<Ticker>>(_response);
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
            return await request.Post(url, requestContent);
        }


        private void BuildUrl(string root, [Optional] string[] args)
        {
            switch (root)
            {
                case "GetAllTickers":
                    url = String.Format(ApiMapping.Roots[root], args[0], Token.value);
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
            if (token.value is null)
            {
                throw new Exception(ConfigError.MissingTokenValue);
            }
            else
            {
                return true;
            }
        }
    }
}