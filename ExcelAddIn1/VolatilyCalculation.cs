using System;
using ExcelAddIn1.PricingCalculation;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1
{
    public static class VolatilyCalculation
    {
        public static void VolatilyMain(Parameters details)
        {
            try
            {
                //affichage du tableau
                details.newWorksheet.Range["B" + details.lastRow].Value = "Volatility Surface";
                details.newWorksheet.Range["B" + details.lastRow].Font.FontStyle = "Bold";
                details.newWorksheet.Range["B" + details.lastRow].Font.Underline = true;
                applyMoneyness(details);
                Grid grid = new Grid(details.mOptionMarketPrice, details.tenors, details.mStrikes);
                details.VolLocale = grid.LocalVolatility(details.mOptionMarketPrice, details.mStrikes, details.tenors,
                    details.r);
                var gv = new GridView(details.newWorksheet, details.mStrikes, details.tenors);
                gv.DisplayGrid(details.lastRow + 1, 3, details.VolLocale);
                gv.DisplayVolSurface("Volatility Surface", details.lastRow + 2, 4);
                details.lastRow += details.mStrikes.Length + 2;
            }
            catch (Exception ex)
            {
                // ignored
            }
        }

        private static void applyMoneyness(Parameters details)
        {
            //application du moneyness, pour ecarter des strikes
            double borneInf;
            int indexBorneInf;
            double borneSup;
            int indexBorneSup;
            borneInf = details.spot * (1 - details.moneyness);
            if (borneInf < 0)
            {
                borneInf = 0;
            }
            borneSup = details.spot * (1 + details.moneyness);
            int i = 0;
            while (details.strikes[i] <= borneInf)
            {
                i++;
            }
            borneInf = details.strikes[i];
            indexBorneInf = i;
            int j = details.strikes.Length - 1;
            while (details.strikes[j] >= borneSup)
            {
                j--;
            }
            borneSup = details.strikes[j];
            indexBorneSup = j;
            details.mStrikes = new double[j - i + 1];
            details.mOptionMarketPrice = new double[j - i + 1, details.tenors.Length];
            for (int k = indexBorneInf; k <= indexBorneSup; k++)
            {
                details.mStrikes[k - indexBorneInf] = details.strikes[k];
                for (int t = 0; t < details.tenors.Length; t++)
                {
                    details.mOptionMarketPrice[k - indexBorneInf, t] = details.optionMarketPrice[k, t];
                }
            }
        }
    }
}