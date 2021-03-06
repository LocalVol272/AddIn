﻿using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

namespace ExcelAddIn1.PricerObjects
{
    public class DataLoader
    {
        public ApiRequest _request;
        protected HttpsRequest request;

        public ApiRequest Request
        {
            get => _request;
            set => _request = value;
        }

        internal Dictionary<string, object> Config { get; set; }


        public void Execute()
        {
            InitRequest();
            //string a = ExecuteRequestAsync().GetAwaiter().GetResult();
            request = new HttpsRequest();
        }

        protected void InitRequest([Optional] Dictionary<string, object> Config)
        {
            _request = MakeRequest(Config);
            _request.BuildRequest();
        }

        protected ApiRequest MakeRequest(Dictionary<string, object> config)
        {
            if (config is null) throw new ArgumentNullException(nameof(config));


            return new ApiRequest(config);
            ;
        }

        private async Task<string> ExecuteRequestAsync()
        {
            // this method will execute the request with proper config
            if (Request.RequestContent.Type == "GET")
                return await Request.Get();
            return await Request.Get();
        }
    }
}