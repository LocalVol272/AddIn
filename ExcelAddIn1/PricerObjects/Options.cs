using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace ExcelAddIn1.PricerObjects
{
    internal class Options : DataLoader, IAuthentification
    {
        private Token _token;
        private string url;
        private string _response;

        private string Reponse
        {
            get => _response;
            set => _response = value;
        }

        public Options()
        {
            _request = new ApiRequest();
        }

        public Options(Token token)
        {
            Token = token;
        }

        public Token Token
        {
            get => _token;
            set => _token = value;
        }

        public Options(Dictionary<string, object> config)
        {
            try
            {
                Config = config;
                Token = GetToken(Config);
                InitRequest(Config);
            }
            catch (Exception _execption)
            {
                throw new Exception(_execption.Message);
            }
        }

        public bool Authentification(Token token)
        {
            if (token.value is null)
                throw new Exception(ConfigError.MissingTokenValue);
            else
                return true;
        }


        public Token GetToken(Dictionary<string, object> config)
        {
            var Token = "Token";

            if (config.ContainsKey(Token)) return new Token(config[Token].ToString());
            throw new Exception(string.Format(ConfigError.MissingKey, Token));
        }


        public Dictionary<string, Dictionary<string, List<Option>>> GetOptions()
        {
            var stack = new StackTrace();
            var root = stack.GetFrame(0).GetMethod().Name;


            var res = new Dictionary<string, Dictionary<string, List<Option>>>();
            Dictionary<string, List<Option>> tempOptionByDates;
            ;
            var ListOptions = new List<Option>();


            foreach (var ticker in (List<string>) Request.RequestContent.Params["Tickers"])
            {
                tempOptionByDates = new Dictionary<string, List<Option>>();

                foreach (var dte in (List<string>) Request.RequestContent.Params["Dates"])
                {
                    string[] args = {ticker, dte};
                    BuilUrl(root, args);
                    _response = ExecuteRequest(url)
                        .GetAwaiter()
                        .GetResult();

                    switch (_response)
                    {
                        case "NotFound":
                            break;
                        default:
                            var str_date = UniversalDateTime.ConvertFromTimestampToString(double.Parse(dte));
                            tempOptionByDates.Add(str_date,
                                FormatOption(
                                    JsonConvert
                                        .DeserializeObject<
                                            Dictionary<string, Dictionary<string, List<Dictionary<string, object>>>>>(
                                            _response), ticker, str_date));
                            break;
                    }
                }

                res.Add(ticker, tempOptionByDates);
            }

            return res;
        }

        private List<Option> FormatOption(
            Dictionary<string, Dictionary<string, List<Dictionary<string, object>>>> option, string ticker, string date)
        {
            string json;
            var _option = new YahooOptionChain();
            var not_available_data = false;
            try
            {
                not_available_data = option["optionChain"]["result"].Count == 0;
            }

            finally
            {
                //not_available_data = JsonConvert.DeserializeObject<List<object>>(JsonConvert.SerializeObject(option["optionChain"]["result"][0]["options"])).Count == 0;
                if (not_available_data)
                {
                    Console.WriteLine(string.Format("On {0} There Are No Available Options For Ticker : {1}", date,
                        ticker));
                }
                else
                {
                    json = JsonConvert.SerializeObject(option["optionChain"]["result"][0]);
                    _option = JsonConvert.DeserializeObject<YahooOptionChain>(json);
                }
            }

            if (Request.RequestContent.Params["Type"].ToString() == "Call")
            {
                if (_option.options is null || _option.options.Count == 0)
                {
                    Console.WriteLine(string.Format("On {0} There Are No Available Options For Ticker : {1}", date,
                        ticker));

                    return null;
                }
                else if (_option.options[0].calls.Count == 0)
                {
                    Console.WriteLine(string.Format("On {0} There Are No Available Options For Ticker : {1}", date,
                        ticker));
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
                    Console.WriteLine(string.Format("On {0} There Are No Available Options For Ticker : {1}", date,
                        ticker));
                    return null;
                }
                else if (_option.options[0].puts.Count == 0)
                {
                    Console.WriteLine(string.Format("On {0} There Are No Available Options For Ticker : {1}", date,
                        ticker));
                    return null;
                }
                else
                {
                    return _option.options[0].puts.ToListOption();
                }
            }
        }


        private void BuilUrl(string root, [Optional] string[] args)
        {
            switch (root)
            {
                case "GetOptions":
                    var type = Request.RequestContent.Params["Type"].ToString();
                    var ticker = args[0];
                    var date = args[1];

                    url = string.Format(ApiMapping.Roots[root], ticker, date);
                    break;
            }
        }

        private async Task<string> ExecuteRequest(string url, [Optional] HttpContent content)
        {
            request = new HttpsRequest();

            if (Request.RequestContent.Type == "GET")
                return await request.Get(url);
            else if (Request.RequestContent.Type == "POST")
                return await request.Post(url, content);
            else
                throw new NotImplementedException("This Request Type Does Not Exist");
        }
    }
}