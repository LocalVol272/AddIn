using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Threading.Tasks;

namespace ExcelAddIn1.PricerObjects
{
    public struct YahooRequest
    {
        public List<string> Tickers { get; set; }
        public string Type { get; set; }
        public string Url { get; set; }

        public HttpContent HttpContent { get; set; }
        public string Response { get; set; }

        public Dictionary<string, object> Params { get; set; }

        public YahooRequest(List<string> RequestTicker, string RequestType, string UrlRequest,
            HttpContent RequestContent = null) : this()
        {
            if (RequestContent is null & (RequestType == "POST"))
                throw new Exception(ConfigError.MissingHttpRequestContent);
            Tickers = RequestTicker;
            Type = RequestType;
            HttpContent = RequestContent;
            Params = null;
        }
    }


    public abstract class HttpRequest
    {
        public abstract void Get(HttpContent Request);
        public abstract void Post(HttpContent Request);
        public abstract void Get(YahooRequest Request);
        public abstract void Post(YahooRequest Request);
        public abstract void Get(object Request);
        public abstract void Post(object Request);
        public abstract Task<string> Get();
        public abstract Task<string> Post();
        public abstract Task<string> Get(string url);
        public abstract Task<string> Post(string url, HttpContent requestContent);
    }
}