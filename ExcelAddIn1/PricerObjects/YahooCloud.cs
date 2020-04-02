using System.Collections.Generic;

namespace ExcelAddIn1.PricerObjects
{
    internal interface IYahooRequest
    {
        Token Token { get; set; }
    }

    internal interface IAuthentification
    {
        Token Token { get; set; }
        Token GetToken(Dictionary<string, object> config);
        bool Authentification(Token token);
    }

    internal interface IYahooResponse
    {
        IEXRequest Read();
    }

    internal interface IYahooApiResponse : IYahooResponse
    {
        bool CheckResponse();
    }

    internal interface IYahooDateFormat
    {
        int Year { get; set; }
        int Month { get; set; }
        int Day { get; set; }
        double ToTimeStamp();
    }
}