using ProjetVolSto.Struct;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Http;
using System.Diagnostics;
using Newtonsoft.Json;
using System.Runtime.InteropServices;

namespace ProjetVolSto.PricerObjects
{
    class Options : DataLoader, IAuthentification
    {
        private Token _token;
        private string url;
        private string _response;
       
        private string Reponse { get => _response; set => _response = value; }
        public Options() => _request = new ApiRequest();
        public Options(Token token) => Token = token;        
        public Token Token { get => _token; set => _token = value; }
        
        public Options(Dictionary<string, object> config)
        {
            try
            {
                this.Config = config;
                this.Token = GetToken(this.Config);
                this.InitRequest(this.Config);
               
            }
            catch (Exception _execption)
            {
                throw new Exception(_execption.Message);

            }
        }

        public bool Authentification(Token token)
        {
            if (token.value is null) { throw new Exception(ConfigError.MissingTokenValue); }
            else { return true; }
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


        public Dictionary<string, Dictionary<string, List<Option>>> GetOptions()
        {
          
            var stack = new StackTrace();
            string root = stack.GetFrame(0).GetMethod().Name;


            Dictionary<string, Dictionary<string, List<Option>>> res = new Dictionary<string, Dictionary<string, List<Option>>>();
            Dictionary<string, List<Option>> tempOptionByDates; ;
            List<Option> ListOptions = new List<Option>();

            

            foreach(string ticker in (List<string>)Request.RequestContent.Params["Tickers"])
            {
                tempOptionByDates = new Dictionary<string, List<Option>>();

                foreach (string dte in (List<string>)Request.RequestContent.Params["Dates"])
                {
                    string[] args = { ticker, dte };
                    BuilUrl(root,args);
                    _response = ExecuteRequest(url)
                        .GetAwaiter()
                        .GetResult();
                     
                    switch (_response)
                    {
                        case "NotFound":
                            break;
                        default:
                            string str_date = UniversalDateTime.ConvertFromTimestampToString(Double.Parse(dte));
                            tempOptionByDates.Add(str_date,
                                FormatOption(JsonConvert.DeserializeObject<Dictionary<string, Dictionary<string, List<Dictionary<string, object>>>>>(_response),ticker, str_date));
                            break ;
                    }


                }
                res.Add(ticker, tempOptionByDates);
            }

            return res;

        }

        private List<Option> FormatOption(Dictionary<string, Dictionary<string, List<Dictionary<string, object>>>> option,string ticker,string date)
        {
            string json;
            YahooOptionChain _option = new YahooOptionChain();
            bool not_available_data=false;
            try
            {
                not_available_data = option["optionChain"]["result"].Count == 0;
                
            }

            finally
            {
                //not_available_data = JsonConvert.DeserializeObject<List<object>>(JsonConvert.SerializeObject(option["optionChain"]["result"][0]["options"])).Count == 0;
                if (not_available_data)
                {
                    Console.WriteLine(String.Format("On {0} There Are No Available Options For Ticker : {1}", date, ticker));

                }
                else
                {
                    json = JsonConvert.SerializeObject(option["optionChain"]["result"][0]);
                    _option = JsonConvert.DeserializeObject<YahooOptionChain>(json);

                }

               
            }

            if (Request.RequestContent.Params["Type"].ToString()=="Call")
            {
                

                    if (_option.options is null || _option.options.Count == 0)
                    {
                        
                            Console.WriteLine(String.Format("On {0} There Are No Available Options For Ticker : {1}", date, ticker));

                        return null;
                    }
                    else if (_option.options[0].calls.Count == 0)
                    {
                    Console.WriteLine(String.Format("On {0} There Are No Available Options For Ticker : {1}", date, ticker));
                    return null;
                    }
                    else
                    {
                        return _option.options[0].calls.ToListOption();
                    }

            }
            else
            {
                if (_option.options is null || _option.options.Count == 0)
                {

                    Console.WriteLine(String.Format("On {0} There Are No Available Options For Ticker : {1}", date, ticker));
                    return null;

                }
                else if(_option.options[0].puts.Count == 0)
                {
                    Console.WriteLine(String.Format("On {0} There Are No Available Options For Ticker : {1}", date, ticker));
                    return null;
                }
                else
                {
                    return _option.options[0].puts.ToListOption();
                }
            }

           

        }



        private void BuilUrl(string root, [Optional]string[] args)
        {
            switch (root)
            {
                
                case "GetOptions":
                    string type = Request.RequestContent.Params["Type"].ToString(); 
                    string ticker = args[0];
                    string date = args[1];

                    url = String.Format(ApiMapping.Roots[root],ticker, date);
                    break;

            }




        }

        private async Task<string> ExecuteRequest(string url,[Optional]HttpContent content)
        {
            request = new HttpsRequest();

            if(Request.RequestContent.Type=="GET")
            {
                return await request.Get(url);
            }
            else if(Request.RequestContent.Type == "POST")
            {
              
                return await request.Post(url, content);
            }
            else{throw new NotImplementedException("This Request Type Does Not Exist");}


        }






    }
}
