﻿using System;
using System.Collections.Generic;
using System.Net;
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

    public class HttpsRequest : HttpRequest
    {
        protected string securedProtocol = "https";

        public override void Get(HttpContent Request)
        {
            throw new NotImplementedException(ApiRequestError.NonImplementedMethod);
        }

        public override void Get(YahooRequest Request)
        {
            throw new NotImplementedException(ApiRequestError.NonImplementedMethod);
        }

        public override void Get(object Request)
        {
            throw new NotImplementedException(ApiRequestError.NonImplementedMethod);
        }

        public override Task<string> Get()
        {
            throw new NotImplementedException(ApiRequestError.NonImplementedMethod);
        }

        public override void Post(HttpContent Request)
        {
            throw new NotImplementedException(ApiRequestError.NonImplementedMethod);
        }

        public override void Post(YahooRequest Request)
        {
            throw new NotImplementedException(ApiRequestError.NonImplementedMethod);
        }

        public override void Post(object Request)
        {
            throw new NotImplementedException(ApiRequestError.NonImplementedMethod);
        }

        public override Task<string> Post()
        {
            throw new NotImplementedException(ApiRequestError.NonImplementedMethod);
        }


        public override async Task<string> Get(string url)
        {
            if (url.Contains(securedProtocol))
            {
                var client = new HttpClient();

                return await ExecuteGet(url, client);
            }

            throw new Exception(HttpRequestError.UnsecuredRequest);
        }

        private static async Task<string> ExecuteGet(string url, HttpClient client)
        {
            try
            {
                var message = await client.GetAsync(url);

                if (message.StatusCode == HttpStatusCode.NotFound)
                {
                    Console.WriteLine("Information Has not Been Found");
                    Console.WriteLine(message);
                    return HttpStatusCode.NotFound.ToString();
                }

                return await message.Content.ReadAsStringAsync();
            }
            catch (Exception _exception)
            {
                Console.WriteLine(_exception);
                return null;
            }
        }

        public override async Task<string> Post(string url, HttpContent requestContent)
        {
            if (url.Contains(securedProtocol))
            {
                var client = new HttpClient();
                return await ExecutePost(url, requestContent, client);
            }

            throw new Exception(HttpRequestError.UnsecuredRequest);
        }

        private static async Task<string> ExecutePost(string url, HttpContent requestContent, HttpClient client)
        {
            try
            {
                var message = await client.PostAsync(url, requestContent);
                if (message.StatusCode == HttpStatusCode.NotFound)
                {
                    Console.WriteLine("Information Has not Been Found");
                    Console.WriteLine(message);
                    return HttpStatusCode.NotFound.ToString();
                }

                return await message.Content.ReadAsStringAsync();
            }
            catch (Exception _exception)
            {
                Console.WriteLine(_exception);
                return null;
            }
        }
    }
}