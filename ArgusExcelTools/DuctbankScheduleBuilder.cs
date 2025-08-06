using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ArgusExcelTools
{
    internal class DuctbankScheduleBuilder
    {
        private readonly Application _excelApp;
        private readonly Workbook _workbook;
        private readonly CreateScheduleHelper _helper;

        public DuctbankScheduleBuilder(Application excelApp, Workbook workbook)
        {
            _excelApp = excelApp;
            _workbook = workbook;
            _helper = new CreateScheduleHelper(_excelApp);
        }

        public void Build()
        {
            Worksheet ductbankSheet = _workbook.Sheets.Add();
            ductbankSheet.Name = "Ductbank";

            //Set column widths and alignment
            double[] columnWidths = { 9, 9, 9, 9, 9, 9, 9 };
            XlHAlign[] columnAlignments =
            {
                XlHAlign.xlHAlignCenter,
                XlHAlign.xlHAlignLeft,
                XlHAlign.xlHAlignLeft,
                XlHAlign.xlHAlignLeft,
                XlHAlign.xlHAlignLeft,
                XlHAlign.xlHAlignLeft,
                XlHAlign.xlHAlignLeft
            };

            for (int i = 0; i < columnWidths.Length; i++)
            {
                int col = i + 2;
                ductbankSheet.Columns[col].ColumnWidth = columnWidths[i];
                ((Range)ductbankSheet.Columns[col]).HorizontalAlignment = columnAlignments[i];
            }

            _helper.ApplyFont(ductbankSheet.Range["B6", "H1000"], "Arial", 10);

            //Header row
            Range headerCell = _helper.MergeAndCenter(ductbankSheet, 2, 2, 2, 8);
            headerCell.Value = "ELECTRICAL DUCTBANK SCHEDULE";
            _helper.FillAndStyleHeader(headerCell, "#E4A444", 18);

            _helper.FillRange(ductbankSheet, 3, 2, 20, 8, "#FFFFFF");

            _helper.MergeAndStyle(ductbankSheet, 4, 2, 5, 2, "EDB-XXX", "#FFFFFF");
            Range r = ductbankSheet.Range["B3", "H20"];
            _helper.ApplyThickBorder(r);

            for(int i = 4; i <= 5; i++)
            {
                for (int j = 3; j <= 6; j++) 
                {
                    Range cell = ductbankSheet.Cells[i, j];
                    _helper.ApplyThickBorder(cell);
                }
            }

            Marshal.ReleaseComObject(ductbankSheet);

        }
    }
}
