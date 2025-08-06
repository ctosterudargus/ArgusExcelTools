using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ArgusExcelTools
{
    internal class CreateScheduleHelper
    {
        private readonly Application _excelApp;

        public CreateScheduleHelper(Application excelApp)
        {
            _excelApp = excelApp;
        }

        public void ApplyFont(Range range, string fontName, int fontSize)
        {
            range.Font.Name = fontName;
            range.Font.Size = fontSize;
        }

        public Range MergeAndCenter(Worksheet sheet, int fromRow, int fromCol, int toRow, int toCol)
        {
            Range range = sheet.Range[
                sheet.Cells[fromRow, fromCol],
                sheet.Cells[toRow, toCol]
            ];

            range.Merge();
            range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            range.VerticalAlignment = XlVAlign.xlVAlignCenter;

            return range;
        }

        public void FillAndStyleHeader(Range cell, string hexColor, int fontSize, bool bold = true, XlHAlign align = XlHAlign.xlHAlignCenter)
        {
            cell.Interior.Color = ColorTranslator.FromHtml(hexColor);
            cell.Font.Bold = bold;
            cell.Font.Size = fontSize;
            cell.HorizontalAlignment = align;
            cell.VerticalAlignment = XlVAlign.xlVAlignCenter;
            cell.Font.Name = "Arial";
            ApplyThickBorder(cell);
        }

        public void FillRange(Worksheet sheet, int fromRow, int fromCol, int toRow, int toCol, string hexColor )
        {
            Range range = sheet.Range[
                sheet.Cells[fromRow, fromCol],
                sheet.Cells[toRow, toCol]
            ];
            range.Interior.Color = ColorTranslator.FromHtml(hexColor);

        }

        public void MergeAndStyle(Worksheet sheet, int fromRow, int fromCol, int toRow, int toCol, string value, string fillHex)
        {
            Range range = MergeAndCenter(sheet, fromRow, fromCol, toRow, toCol);
            range.Value = value;
            range.Interior.Color = ColorTranslator.FromHtml(fillHex);
            range.Font.Bold = true;
            range.Font.Name = "Arial";
            range.Font.Size = 10;
            ApplyThickBorder(range);
        }

        public void StyleSingleCell(Worksheet sheet, int row, int col, string value, string fillHex)
        {
            Range cell = sheet.Cells[row, col];
            cell.Value = value;
            cell.Interior.Color = ColorTranslator.FromHtml(fillHex);
            cell.Font.Bold = true;
            cell.Font.Name = "Arial";
            cell.Font.Size = 10;
            cell.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            cell.VerticalAlignment = XlVAlign.xlVAlignCenter;
            ApplyThickBorder(cell);
        }

        public void ApplyThickBorder(Range range)
        {
            range.Borders[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlMedium;
            range.Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlMedium;
            range.Borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlMedium;
            range.Borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlMedium;
        }

        public void ApplyThinTableBorder(Range cell)
        {
            cell.Borders[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlThin;
            cell.Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThin;
            cell.Borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlThin;
            cell.Borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlThin;
        }
    }
}
