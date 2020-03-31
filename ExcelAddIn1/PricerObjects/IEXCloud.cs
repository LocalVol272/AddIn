using System.Collections.Generic;
using ProjetVolSto.Struct;

namespace ProjetVolSto.PricerObjects
{
    interface IEXCloudRequest
    {
        Token Token { get; set; }
    }

    interface IAuthentification
    {
        Token Token { get; set; }
        Token GetToken(Dictionary<string, object> config);
        bool Authentification(Token token);
    }

    interface IEXResponse
    {
        IEXRequest Read();
    }

    interface IEXCloudResponse : IEXResponse
    {
        bool CheckResponse();
    }

    interface IEXDate
    {
        string Year { get; set; }
        string Month { get; set; }
        string Day { get; set; }
        string Format();
    }
}