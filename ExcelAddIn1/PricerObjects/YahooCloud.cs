using System.Collections.Generic;

namespace ExcelAddIn1.PricerObjects
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
        YahooRequest Read();
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
