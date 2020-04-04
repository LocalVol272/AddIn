using System;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1
{
    public class NewWorksheet
    {
        private static Worksheet _newWorksheet;

        public static void Creation(Worksheet newWorksheet)
        {
            //creation d'une nouvelle WS
            _newWorksheet = newWorksheet;
            string creationSheetDate = DateTime.Now.ToString("HH:mm:ss");
            var nameSheet = "LocalVolatilityWS_" + creationSheetDate.Replace(":", "_");
            try
            {
                _newWorksheet.Name = nameSheet;
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message + " Merci d'essayer à nouveau.");
            }
            
            VisualizeData();
        }

        private static void VisualizeData()
        {
            //visualisation des datas
            SetText();
            SetColor();
            _newWorksheet.Columns.AutoFit();
        }

        private static void SetColor()
        {
            _newWorksheet.Range["A1", "A3"].Font.ColorIndex = 1;
            _newWorksheet.Range["A1", "A3"].Font.Size = 11;
            _newWorksheet.Range["A1", "A3"].Font.FontStyle = "Bold";
            _newWorksheet.Range["A1", "A3"].Interior.Color = 14599344;
            _newWorksheet.Range["A1", "B3"].Borders.LineStyle = XlLineStyle.xlContinuous;
            _newWorksheet.Range["A1", "B3"].Borders.Weight = 2d;
        }

        private static void SetText()
        {
            _newWorksheet.Range["A1"].Value = "Ticker";
            _newWorksheet.Range["A2"].Value = "Underlying Price";
            _newWorksheet.Range["A3"].Value = "Option Type";
        }
    }
}