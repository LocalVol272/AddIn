namespace ProjetVolSto.Struct
{
    public struct Token
    {
        public string value;

        public Token(string _value) => value = _value;

    }

    public struct IEXApiAdress
    {
        public const string Url = "https://sandbox.iexapis.com/stable/";
    }






}