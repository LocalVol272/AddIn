using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1
{
    public class GridView
    {
        private  Worksheet _Worksheet;
        private int[] _strikes;
        private double[] _tenors; 
        public GridView(Worksheet ws, int [] strike, double [] tenors)
        {
            _Worksheet = ws;
            _strikes = strike;
            _tenors = tenors;
        }

        public void DisplayGrid(int rowCoordinate, int columnCoordinate, double[,] data)
        {
            DisplayStrikes(rowCoordinate, columnCoordinate);
            DisplayTenors(rowCoordinate, columnCoordinate);
            DisplayInsideGrid(rowCoordinate, columnCoordinate, data);
            _Worksheet.Columns.AutoFit();
        }

        private void DisplayStrikes(int rowCoordinate, int columnCoordinate)
        {
            int strikes_coord = rowCoordinate + _strikes.Length - 1;
            Range strikeRange = _Worksheet.Range[_Worksheet.Cells[rowCoordinate + 2, columnCoordinate], _Worksheet.Cells[strikes_coord + 2, columnCoordinate]];
            strikeRange.Merge();
            strikeRange.Orientation = 90;
            strikeRange.Value = "Strikes";
            strikeRange.Font.FontStyle = "Bold";
            strikeRange.Interior.Color = 14599344;
            strikeRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            strikeRange.VerticalAlignment = XlVAlign.xlVAlignCenter;
            int curr_coord;
            for (int i = 0; i < _strikes.Length; i++)
            {
                curr_coord = rowCoordinate + i +2;
                _Worksheet.Cells[curr_coord, columnCoordinate +1].Value = _strikes[i];
                _Worksheet.Cells[curr_coord, columnCoordinate + 1].Interior.Color = 13882323;
            }
            Range _cell1 = _Worksheet.Cells[rowCoordinate + 2, columnCoordinate];
            Range _cell2 = _Worksheet.Cells[rowCoordinate + 2, columnCoordinate+1];
            Range _cell3 = _Worksheet.Cells[strikes_coord + 2, columnCoordinate + 1];
            _Worksheet.Range[_cell2, _cell3].Borders.LineStyle = XlLineStyle.xlContinuous;
            _Worksheet.Range[_cell2, _cell3].Borders.Weight = 2d;
            _Worksheet.Range[_cell1, _cell3].Borders.LineStyle = XlLineStyle.xlContinuous;
            _Worksheet.Range[_cell1, _cell3].Borders.Weight = 3d;


        }
        private void DisplayTenors(int rowCoordinate, int columnCoordinate)
        {
            int tenors_coord = columnCoordinate + _tenors.Length - 1;
            Range tenorRange = _Worksheet.Range[_Worksheet.Cells[rowCoordinate, columnCoordinate + 2], _Worksheet.Cells[rowCoordinate, tenors_coord + 2]];
            tenorRange.Merge();
            tenorRange.Value = "Tenors";
            tenorRange.Font.FontStyle = "Bold";
            tenorRange.Interior.Color = 14599344;
            tenorRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            tenorRange.VerticalAlignment = XlVAlign.xlVAlignCenter;
            int curr_coord;
            for (int i = 0; i < _tenors.Length; i++)
            {
                curr_coord = columnCoordinate + i;
                _Worksheet.Cells[rowCoordinate + 1, curr_coord +2].Value = _tenors[i];
                _Worksheet.Cells[rowCoordinate + 1, curr_coord + 2].Interior.Color = 13882323;
            }
            Range _cell1 = _Worksheet.Cells[rowCoordinate, columnCoordinate+2];
            Range _cell2 = _Worksheet.Cells[rowCoordinate + 1, columnCoordinate + 2];
            Range _cell3 = _Worksheet.Cells[ rowCoordinate + 1, tenors_coord +2];
            _Worksheet.Range[_cell2, _cell3].Borders.LineStyle = XlLineStyle.xlContinuous;
            _Worksheet.Range[_cell2, _cell3].Borders.Weight = 2d;
            _Worksheet.Range[_cell1, _cell3].Borders.LineStyle = XlLineStyle.xlContinuous;
            _Worksheet.Range[_cell1, _cell3].Borders.Weight = 3d;
            _Worksheet.Range[_cell2, _cell3].HorizontalAlignment = XlHAlign.xlHAlignCenter;
        }

        private void DisplayInsideGrid(int rowCoordinate, int columnCoordinate, double[,] data)
        {
            for (int i = 0; i < _strikes.Length; i++)
            {
                for (int j = 0; j < _tenors.Length; j++)
                {
                    _Worksheet.Cells[rowCoordinate + i +2 , columnCoordinate + j + 2].Value = data[i, j];
                }
            }
            Range _cell1 = _Worksheet.Cells[rowCoordinate+2, columnCoordinate + 2];
            Range _cell2 = _Worksheet.Cells[rowCoordinate + _strikes.Length + 1 , columnCoordinate + _tenors.Length + 1];
            _Worksheet.Range[_cell1,_cell2].Borders.LineStyle = XlLineStyle.xlContinuous;
            _Worksheet.Range[_cell1, _cell2].Borders.Weight = 2d;
            _Worksheet.Range[_cell1, _cell2].HorizontalAlignment = XlHAlign.xlHAlignCenter;
        }

        
    }
}
