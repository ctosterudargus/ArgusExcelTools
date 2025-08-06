using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ArgusExcelTools
{
    internal class CableTrayScheduleBuilder
    {
        private readonly Application _excelApp;
        private readonly Workbook _workbook;
        private readonly CreateScheduleHelper _helper;

        public CableTrayScheduleBuilder(Application excelApp, Workbook workbook)
        {
            _excelApp = excelApp;
            _workbook = workbook;
            _helper = new CreateScheduleHelper(_excelApp);
        }

        public void Build()
        {
            Worksheet cableTraySheet = _workbook.Sheets.Add();
            cableTraySheet.Name = "Cable Tray";

            // Set column widths and alignment
            double[] columnWidths = { 11, 11, 14, 9, 9, 9, 14, 9, 9, 9, 14, 9, 9, 9 };
            XlHAlign[] columnAlignments = {
                XlHAlign.xlHAlignCenter,
                XlHAlign.xlHAlignCenter,
                XlHAlign.xlHAlignCenter,
                XlHAlign.xlHAlignLeft,
                XlHAlign.xlHAlignLeft,
                XlHAlign.xlHAlignLeft,
                XlHAlign.xlHAlignCenter,
                XlHAlign.xlHAlignLeft,
                XlHAlign.xlHAlignLeft,
                XlHAlign.xlHAlignLeft,
                XlHAlign.xlHAlignCenter,
                XlHAlign.xlHAlignLeft,
                XlHAlign.xlHAlignLeft,
                XlHAlign.xlHAlignLeft
            };

            for (int i = 0; i < columnWidths.Length; i++)
            {
                int col = i + 2;
                cableTraySheet.Columns[col].ColumnWidth = columnWidths[i];
                ((Range)cableTraySheet.Columns[col]).HorizontalAlignment = columnAlignments[i];
            }

            // Set font
            _helper.ApplyFont(cableTraySheet.Range["B6", "Q1000"], "Arial", 10);

            // Header row
            Range headerCell = _helper.MergeAndCenter(cableTraySheet, 2, 2, 2, 15);
            headerCell.Value = "ELECTRICAL CABLE TRAY SCHEDULE";
            _helper.FillAndStyleHeader(headerCell, "#E4A444", 18);

            // Column headers (gray)
            string gray = "#D9D9D9";

            _helper.MergeAndStyle(cableTraySheet, 3, 2, 4, 2, "TRAY ID", gray);
            _helper.MergeAndStyle(cableTraySheet, 3, 3, 4, 3, "TRAY SIZE", gray);
            _helper.MergeAndStyle(cableTraySheet, 3, 4, 4, 4, "POWER SIZE", gray);
            _helper.MergeAndStyle(cableTraySheet, 3, 5, 4, 7, "POWER CABLES", gray);
            _helper.MergeAndStyle(cableTraySheet, 3, 8, 4, 8, "CONTROL SIZE", gray);
            _helper.MergeAndStyle(cableTraySheet, 3, 9, 4, 11, "CONTROL CABLES", gray);
            _helper.MergeAndStyle(cableTraySheet, 3, 12, 4, 12, "DATA SIZE", gray);
            _helper.MergeAndStyle(cableTraySheet, 3, 13, 4, 15, "DATA CABLES", gray);


            _helper.MergeAndStyle(cableTraySheet, 5, 2, 16, 2, "ECT-XXX", gray);
            _helper.MergeAndStyle(cableTraySheet, 17, 2, 28, 2, "ECT-XXX", gray);
            _helper.FillRange(cableTraySheet, 6, 3, 28, 15, "#FFFFFF");

            /*
            

            // Example data row
            Range rowLabel = _helper.MergeAndCenter(cableTraySheet, 6, 2, 6, 3);
            rowLabel.Value = "UTILITY POWER";
            rowLabel.Font.Bold = true;

            // Table data border styling
            for (int row = 6; row <= 30; row++)
            {
                for (int col = 2; col <= 10; col++)
                {
                    Range cell = cableTraySheet.Cells[row, col];
                    _helper.ApplyThinTableBorder(cell);

                    if (col == 2)
                        cell.Borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlMedium;

                    if (col == 10)
                        cell.Borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlMedium;

                    if (row == 30)
                        cell.Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlMedium;
                }
            } */

            Marshal.ReleaseComObject(cableTraySheet);
        }
    }
}
