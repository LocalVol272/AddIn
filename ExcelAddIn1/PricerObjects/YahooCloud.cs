using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ProjetVolSto.Struct;
using System.Net.Http;
namespace ProjetVolSto.PricerObjects
{
   

    interface IYahooRequest
    {
        Token Token { get; set; }
       

    }

    interface IAuthentification
    {
        Token Token { get; set; }
        Token GetToken(Dictionary<string, object> config);
        bool Authentification(Token token);
    }

    interface IYahooResponse
    {
        IEXRequest Read();
    }

    interface IYahooApiResponse: IYahooResponse
    {
       bool CheckResponse();
        
    }

    interface IYahooDateFormat
    {
        int Year { get; set; }
        int Month{ get; set; }
        int Day { get; set; }
        double ToTimeStamp();
    }




}
