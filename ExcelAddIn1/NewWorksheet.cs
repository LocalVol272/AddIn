using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1
{
    public class NewWorksheet
    {
        private static Worksheet _newWorksheet;
        private static string _creationSheetDate;

        public static void Creation(Worksheet newWorksheet, string creationSheetDate)
        {
            _newWorksheet = newWorksheet;
            _creationSheetDate = creationSheetDate;
            var nameSheet = "LocalVolatilityWS_" + _creationSheetDate.Replace(":", "_");
            _newWorksheet.Name = nameSheet;
            VisualizeData();
        }

        private static void VisualizeData()
        {
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