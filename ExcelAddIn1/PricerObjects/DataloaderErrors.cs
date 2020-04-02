namespace ExcelAddIn1.PricerObjects
{
    public static class ConfigError
    {
        public const string MissingTokenValue = "Token Key Do Not Exist In Config";
        public const string MissingKey = "The {0} Is Not Refer In The Dictionary";
        public const string MissingHttpRequestContent = "This Request Cannot Be Executed : Missing _RequestContent";
    }

    public static class DataLoaderError
    {
        public const string DateFormatError = "Date Format Should Be YYYYMMDD or YYYYMM";
    }

    public static class ApiRequestError
    {
        public const string NonImplementedMethod = "Non Implemented Method";
    }
}