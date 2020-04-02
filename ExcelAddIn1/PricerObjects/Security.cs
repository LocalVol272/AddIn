using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Threading.Tasks;

namespace ExcelAddIn1.PricerObjects
{
    internal class ApiRequest : HttpRequest, IYahooRequest, IAuthentification
    {
        protected Token token;
        private IEXRequest request;
        private Dictionary<string, object> config;

        Token IYahooRequest.Token
        {
            get => token;
            set => token = value;
        }

        Token IAuthentification.Token
        {
            get => token;
            set => token = value;
        }

        public IEXRequest RequestContent
        {
            get => request;
            set => request = value;
        }

        public override void Get(IEXRequest Request)
        {
            throw new NotImplementedException(ApiRequestError.NonImplementedMethod);
        }

        public override void Post(IEXRequest Request)
        {
            throw new NotImplementedException(ApiRequestError.NonImplementedMethod);
        }

        public override void Get(HttpContent Request)
        {
            throw new NotImplementedException(ApiRequestError.NonImplementedMethod);
        }

        public override void Post(HttpContent Request)
        {
            throw new NotImplementedException(ApiRequestError.NonImplementedMethod);
        }

        public override void Get(object Request)
        {
            throw new NotImplementedException(ApiRequestError.NonImplementedMethod);
        }

        public override void Post(object Request)
        {
            throw new NotImplementedException(ApiRequestError.NonImplementedMethod);
        }

        public override Task<string> Get(string url)
        {
            throw new NotImplementedException(ApiRequestError.NonImplementedMethod);
        }

        public override Task<string> Post(string url, HttpContent requestContent)
        {
            throw new NotImplementedException(ApiRequestError.NonImplementedMethod);
        }

        public bool Authentification(Token token)
        {
            throw new NotImplementedException();
        }

        public ApiRequest()
        {
            ;
        }

        public ApiRequest(Dictionary<string, object> config)
        {
            this.config = config;
            token = GetToken(config);
        }


        public Token GetToken(Dictionary<string, object> config)
        {
            var Token = "Token";

            if (config.ContainsKey(Token)) return new Token(config[Token].ToString());
            throw new Exception(string.Format(ConfigError.MissingKey, Token));
        }

        public void BuildRequest()
        {
            SetRequestType();
            SetParams();
            SetTickers();
            UnWrapParams();
            BuildUrl();
            RequestContent = request;
        }

        private void SetTickers()
        {
            var Tickers = "Tickers";

            if (request.Params.ContainsKey(Tickers))
                request.Params[Tickers] = (List<string>) request.Params[Tickers];
            else
                throw new Exception(string.Format(ConfigError.MissingKey, Tickers));
            ;
        }

        private void SetRequestType()
        {
            var Type = "Type";

            if (config.ContainsKey(Type))
                request.Type = config[Type].ToString();
            else
                throw new Exception(string.Format(ConfigError.MissingKey, Type));
            ;
        }

        private void BuildUrl()
        {
        }


        private void SetParams()
        {
            var Params = "Params";

            if (config.ContainsKey(Params))
            {
                if (config[Params].GetType() == typeof(Dictionary<string, object>))
                    request.Params = (Dictionary<string, object>) config[Params];
            }
            else
            {
                throw new Exception(string.Format(ConfigError.MissingKey, Params));
            }

            ;
        }


        private void UnWrapParams()
        {
            request.Params["Dates"] = (List<string>) request.Params["Dates"];
            SetDateFormat();
            SetProductType();
        }

        private void SetDateFormat()
        {
            var DateList = new List<string>();

            foreach (var dte in (List<string>) request.Params["Dates"])
            {
                IYahooDateFormat iEXDate = new Date(dte);

                DateList.Add(iEXDate.ToTimeStamp().ToString());
            }

            request.Params["Dates"] = DateList;
        }

        private void SetProductType()
        {
            var productType = "ProductType";
            if (request.Params.ContainsKey(productType))
            {
                var separator = '/';
                var product_type = request.Params[productType].ToString();
                var args = product_type.Split(separator);
                request.Params.Add("Product", args[0]);
                request.Params.Add("Type", args[1]);
            }
            else
            {
                throw new Exception(string.Format(ConfigError.MissingKey, productType));
            }
        }


        public override async Task<string> Post()
        {
            var client = new HttpClient();

            try
            {
                var message = await client.PostAsync(RequestContent.Url, RequestContent.HttpContent);
                return await message.Content.ReadAsStringAsync();
            }
            catch (Exception _exception)
            {
                Console.WriteLine(_exception);
            }

            return null;
        }

        public override async Task<string> Get()
        {
            var client = new HttpClient();
            try
            {
                var message =
                    await client.GetAsync(
                        "https://sandbox.iexapis.com/stable/stock/aapl/options/202001?token=Tsk_bbe66f58b6d149f59a9af4eb83bfc7f5");

                Console.WriteLine(message.Content.ToString());
                return await message.Content.ReadAsStringAsync();
            }
            catch (Exception _exception)
            {
                Console.WriteLine(_exception);
            }

            return null;
        }
    }
}