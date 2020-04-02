using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1
{
    public class GridView
    {
        private readonly int[] _strikes;
        private readonly double[] _tenors;
        private readonly Worksheet _worksheet;

        public GridView(Worksheet ws, int[] strike, double[] tenors)
        {
            _worksheet = ws;
            _strikes = strike;
            _tenors = tenors;
        }

        public void DisplayGrid(int rowCoordinate, int columnCoordinate, double[,] data)
        {
            DisplayStrikes(rowCoordinate, columnCoordinate);
            DisplayTenors(rowCoordinate, columnCoordinate);
            DisplayInsideGrid(rowCoordinate, columnCoordinate, data);
            _worksheet.Columns.AutoFit();
        }

        private void DisplayStrikes(int rowCoordinate, int columnCoordinate)
        {
            var strikesCoord = rowCoordinate + _strikes.Length - 1;
            var strikeRange = _worksheet.Range[_worksheet.Cells[rowCoordinate + 2, columnCoordinate],
                _worksheet.Cells[strikesCoord + 2, columnCoordinate]];
            strikeRange.Merge();
            strikeRange.Orientation = 90;
            InsideFont(strikeRange, "Strikes");
            for (var i = 0; i < _strikes.Length; i++)
            {
                var currCoord = rowCoordinate + i + 2;
                _worksheet.Cells[currCoord, columnCoordinate + 1].Value = _strikes[i];
                _worksheet.Cells[currCoord, columnCoordinate + 1].Interior.Color = 13882323;
            }

            Range cell1 = _worksheet.Cells[rowCoordinate + 2, columnCoordinate];
            Range cell2 = _worksheet.Cells[rowCoordinate + 2, columnCoordinate + 1];
            Range cell3 = _worksheet.Cells[strikesCoord + 2, columnCoordinate + 1];
            BordersStyleTwo(cell2, cell3, cell1);
        }

        private void DisplayTenors(int rowCoordinate, int columnCoordinate)
        {
            var tenorsCoord = columnCoordinate + _tenors.Length - 1;
            var tenorRange = _worksheet.Range[_worksheet.Cells[rowCoordinate, columnCoordinate + 2],
                _worksheet.Cells[rowCoordinate, tenorsCoord + 2]];
            tenorRange.Merge();
            InsideFont(tenorRange, "Tenors");
            for (var i = 0; i < _tenors.Length; i++)
            {
                var currCoord = columnCoordinate + i;
                _worksheet.Cells[rowCoordinate + 1, currCoord + 2].Value = _tenors[i];
                _worksheet.Cells[rowCoordinate + 1, currCoord + 2].Interior.Color = 13882323;
            }

            Range cell1 = _worksheet.Cells[rowCoordinate, columnCoordinate + 2];
            Range cell2 = _worksheet.Cells[rowCoordinate + 1, columnCoordinate + 2];
            Range cell3 = _worksheet.Cells[rowCoordinate + 1, tenorsCoord + 2];
            BordersStyleTwo(cell2, cell3, cell1);
        }

        private void BordersStyleTwo(Range cell2, Range cell3, Range cell1)
        {
            _worksheet.Range[cell2, cell3].Borders.LineStyle = XlLineStyle.xlContinuous;
            _worksheet.Range[cell2, cell3].Borders.Weight = 2d;
            _worksheet.Range[cell1, cell3].Borders.LineStyle = XlLineStyle.xlContinuous;
            _worksheet.Range[cell1, cell3].Borders.Weight = 3d;
            _worksheet.Range[cell2, cell3].HorizontalAlignment = XlHAlign.xlHAlignCenter;
        }

        private static void InsideFont(Range tenorRange, string type)
        {
            tenorRange.Value = type;
            tenorRange.Font.FontStyle = "Bold";
            tenorRange.Interior.Color = 14599344;
            tenorRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            tenorRange.VerticalAlignment = XlVAlign.xlVAlignCenter;
        }

        private void DisplayInsideGrid(int rowCoordinate, int columnCoordinate, double[,] data)
        {
            for (var i = 0; i < _strikes.Length; i++)
            for (var j = 0; j < _tenors.Length; j++)
                _worksheet.Cells[rowCoordinate + i + 2, columnCoordinate + j + 2].Value = data[i, j];
            Range cell1 = _worksheet.Cells[rowCoordinate + 2, columnCoordinate + 2];
            Range cell2 = _worksheet.Cells[rowCoordinate + _strikes.Length + 1, columnCoordinate + _tenors.Length + 1];
            BordersStyleOne(cell1, cell2);
        }

        private void BordersStyleOne(Range cell1, Range cell2)
        {
            _worksheet.Range[cell1, cell2].Borders.LineStyle = XlLineStyle.xlContinuous;
            _worksheet.Range[cell1, cell2].Borders.Weight = 2d;
            _worksheet.Range[cell1, cell2].HorizontalAlignment = XlHAlign.xlHAlignCenter;
        }

        public void DisplayVolSurface(string title, int rowIndexDataSource, int columnIndexDataSource)
        {
            Shape VolSurfShape = _worksheet.Shapes.AddChart2(Width: 600, Height: 300);
            Chart VolSurf = VolSurfShape.Chart;
            VolSurf.HasTitle = true;
            VolSurf.ChartTitle.Text = title;
            Range _cell1 = _worksheet.Cells[rowIndexDataSource, columnIndexDataSource];
            Range _cell2 = _worksheet.Cells[rowIndexDataSource + _strikes.Length, columnIndexDataSource + _tenors.Length];
            VolSurf.SetSourceData(_worksheet.Range[_cell1, _cell2]);
            VolSurf.ChartType = XlChartType.xlSurface;
            VolSurf.ChartStyle = 311;
            VolSurf.ChartColor = 21;
            VolSurf.Location(XlChartLocation.xlLocationAsNewSheet, "Volatility Surface");

        }
    }
}