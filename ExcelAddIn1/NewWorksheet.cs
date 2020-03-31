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
            var nameSheet = "OptionPricing_" + _creationSheetDate.Replace(":", "_");
            _newWorksheet.Name = nameSheet;
            VisualizeData();
            //"Ok";
        }

        private static void VisualizeData()
        {
            SetText();
            SetColor();

            _newWorksheet.Range["A1", "Z1000"].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            _newWorksheet.Range["A1", "Z1000"].VerticalAlignment = XlHAlign.xlHAlignCenter;
            _newWorksheet.Columns.AutoFit();
        }

        private static void SetColor()
        {
            MakeBorders(_newWorksheet.Range["A1", "Z1000"], "White");

            _newWorksheet.Range["A1"].Font.ColorIndex = 46;
            _newWorksheet.Range["A1"].Font.Size = 16;
            _newWorksheet.Range["A1", "F3"].Font.FontStyle = "Bold";
            _newWorksheet.Range["F1"].Font.ColorIndex = 3;
            _newWorksheet.Range["E1"].Font.ColorIndex = 1;

            _newWorksheet.Range["A3", "B6"].Interior.ColorIndex = 43;
            _newWorksheet.Range["A8", "B13"].Interior.ColorIndex = 37;

            _newWorksheet.Range["A8"].Font.FontStyle = "Bold";
            _newWorksheet.Range["B4", "B6"].Interior.ColorIndex = 44;
            _newWorksheet.Range["B14", "B15"].Interior.ColorIndex = 3;
            _newWorksheet.Range["B9", "B15"].Interior.ColorIndex = 44;
            _newWorksheet.Range["A14", "A15"].Interior.ColorIndex = 3;

            MakeBorders(_newWorksheet.Range["A3", "B6"], "Black");
            MakeBorders(_newWorksheet.Range["A8", "B15"], "Black");
        }

        private static void SetText()
        {
            _newWorksheet.Range["A1"].Value = "Local volatility";
            _newWorksheet.Range["E1"].Value = "Creation time :";
            _newWorksheet.Range["F1"].Value = _creationSheetDate;
            _newWorksheet.Range["A1", "D1"].Merge();

            _newWorksheet.Range["A3"].Value = "Input";
            _newWorksheet.Range["A3", "B3"].Merge();
            _newWorksheet.Range["A4"].Value = "Ticker";
            _newWorksheet.Range["A5"].Value = "Maturity";
            _newWorksheet.Range["A6"].Value = "Option";

            _newWorksheet.Range["A8"].Value = "Output";
            _newWorksheet.Range["A8", "B8"].Merge();
            _newWorksheet.Range["A9"].Value = "Bid";
            _newWorksheet.Range["A10"].Value = "Ask";
            _newWorksheet.Range["A11"].Value = "Strike";
            _newWorksheet.Range["A12"].Value = "Closing Price";
            _newWorksheet.Range["A13"].Value = "Strike";
            _newWorksheet.Range["A14"].Value = "Volatility";
            _newWorksheet.Range["A15"].Value = "Option Price";

        }

        private static void MakeBorders(Range range, string color)
        {
            range.Borders.LineStyle = XlLineStyle.xlContinuous;
            switch (color)
            {
                case "White":
                    range.Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 255, 255));
                    break;
                default:
                    range.Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(0, 0, 0));
                    break;
            }
        }
    }
}