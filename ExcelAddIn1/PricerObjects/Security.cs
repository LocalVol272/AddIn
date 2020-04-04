using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Threading.Tasks;

namespace ExcelAddIn1.PricerObjects
{
    public class ApiRequest : HttpRequest, IYahooRequest,IAuthentification
    {
        protected Token token;
        private YahooRequest request;
        private Dictionary<string, object> config;
        Token IYahooRequest.Token { get => this.token; set => this.token = value; }
        Token IAuthentification.Token { get => this.token; set => this.token =value; }
        public YahooRequest RequestContent { get => request; set => request = value; }
        public override void Get(YahooRequest Request) => throw new NotImplementedException(ApiRequestError.NonImplementedMethod);
        public override void Post(YahooRequest Request) => throw new NotImplementedException(ApiRequestError.NonImplementedMethod);
        public override void Get(HttpContent Request) => throw new NotImplementedException(ApiRequestError.NonImplementedMethod);
        public override void Post(HttpContent Request) => throw new NotImplementedException(ApiRequestError.NonImplementedMethod);
        public override void Get(object Request) => throw new NotImplementedException(ApiRequestError.NonImplementedMethod);
        public override void Post(object Request) => throw new NotImplementedException(ApiRequestError.NonImplementedMethod);
        public override Task<string> Get(string url) => throw new NotImplementedException(ApiRequestError.NonImplementedMethod);
        public override Task<string> Post(string url, HttpContent requestContent) => throw new NotImplementedException(ApiRequestError.NonImplementedMethod);
        public bool Authentification(Token token) => throw new NotImplementedException();
        public ApiRequest() {; }

        public ApiRequest(Dictionary<string, object> config)
        {
            this.config = config;
            this.token = this.GetToken(config);
            
        }

        

       





        public Token GetToken(Dictionary<string, object> config)
        {
            string Token = "Token";

            if (config.ContainsKey(Token))
            {
                return new Token(config[Token].ToString()) ;
            }
            throw new Exception(String.Format(ConfigError.MissingKey, Token));
        }

        public void BuildRequest()
        {

            
            this.SetRequestType();
            this.SetParams();
            this.SetTickers();
            this.UnWrapParams();
            this.BuildUrl();
            this.RequestContent = request;
            
        }

        private void SetTickers()
        {
            string Tickers = "Tickers";

            if (this.request.Params.ContainsKey(Tickers))
            {


                this.request.Params[Tickers]= (List<string>)this.request.Params[Tickers];
            
            
            
            }
            else { throw new Exception(String.Format(ConfigError.MissingKey, Tickers));};


      
        }

        private void SetRequestType()
        {
            string Type = "Type";

            if (config.ContainsKey(Type))
            {
                this.request.Type = this.config[Type].ToString();
            }
            else { throw new Exception(String.Format(ConfigError.MissingKey, Type)); };

        }

        private void BuildUrl()
        {

        }
       

        private void SetParams()
        {
            string Params = "Params";

            if (config.ContainsKey(Params))
            {
                if (this.config[Params].GetType() == typeof(Dictionary<string, object>))
                {


                    this.request.Params = (Dictionary<string, object>)this.config[Params];
                    
                }
            }
            else { throw new Exception(String.Format(ConfigError.MissingKey, Params)); };
        }


        private void UnWrapParams()
        {

            
            this.request.Params["Dates"] = (List<string>)this.request.Params["Dates"];
            SetDateFormat();
            SetProductType();

        }
        private void SetDateFormat()
        {
            var DateList = new List<string>();
            
            foreach(string dte in (List<string>)this.request.Params["Dates"])
            {
  
                IYahooDateFormat iEXDate = new Date(dte);

                DateList.Add(iEXDate.ToTimeStamp().ToString());
            }

            this.request.Params["Dates"] = DateList;

        }

        private void SetProductType()
        {
            string productType = "ProductType";
            if (this.request.Params.ContainsKey(productType))
            {
                char separator = '/';
                string product_type = this.request.Params[productType].ToString();
                string[] args = product_type.Split(separator);
                this.request.Params.Add("Product",args[0]);
                this.request.Params.Add("Type", args[1]);

            }
            else { throw new Exception(String.Format(ConfigError.MissingKey, productType));}
        }




        public override async Task<string> Post()
        {
            HttpClient client = new HttpClient();

            try
            {
                HttpResponseMessage message = await client.PostAsync(RequestContent.Url, RequestContent.HttpContent);
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
            HttpClient client = new HttpClient();
            try
            {
                HttpResponseMessage message = await client.GetAsync("https://sandbox.iexapis.com/stable/stock/aapl/options/202001?token=Tsk_bbe66f58b6d149f59a9af4eb83bfc7f5");
                
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


