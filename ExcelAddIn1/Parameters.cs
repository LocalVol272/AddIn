using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1
{
    public class Parameters
    {
        public int lastRow;
        public Worksheet newWorksheet;
        public double spot;
        public string type;
        public double r;
        public double moneyness = 2.5;
        public double[] strikes;
        public double[] mStrikes;
        public double[] tenors;
        public double[,] optionMarketPrice;
        public double[,] mOptionMarketPrice;
        public double[,] VolLocale;
    }
}