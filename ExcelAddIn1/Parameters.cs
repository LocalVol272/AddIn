using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1
{
    public class Parameters
    {
        public int lastRow;
        public double moneyness = 2.5;
        public double[,] mOptionMarketPrice;
        public double[] mStrikes;
        public Worksheet newWorksheet;
        public double[,] optionMarketPrice;
        public double r;
        public double spot;
        public double[] strikes;
        public double[] tenors;
        public string type;
        public double[,] VolLocale;
    }
}