using System;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;

namespace ExcelAddIn1.PricerObjects
{
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