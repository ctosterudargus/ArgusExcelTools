using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ArgusExcelTools
{
    internal class CableScheduleBuilder
    {
        private readonly Application _excelApp;
        private readonly Workbook _workbook;
        private readonly CreateScheduleHelper _helper;

        public CableScheduleBuilder(Application excelApp, Workbook workbook)
        {
            _excelApp = excelApp;
            _workbook = workbook;
            _helper = new CreateScheduleHelper(_excelApp);
        }

        public void Build()
        {
            Worksheet cablesSheet = _workbook.Sheets.Add();
            cablesSheet.Name = "Cables";

            // Set column widths and alignment
            double[] columnWidths = { 3, 11, 34, 34, 6, 14, 11, 6, 78, 69, 17, 4, 4, 4, 14, 70 };
            XlHAlign[] columnAlignments = {
                XlHAlign.xlHAlignLeft,
                XlHAlign.xlHAlignLeft,
                XlHAlign.xlHAlignLeft,
                XlHAlign.xlHAlignLeft,
                XlHAlign.xlHAlignCenter,
                XlHAlign.xlHAlignLeft,
                XlHAlign.xlHAlignCenter,
                XlHAlign.xlHAlignCenter,
                XlHAlign.xlHAlignLeft,
                XlHAlign.xlHAlignLeft,
                XlHAlign.xlHAlignCenter,
                XlHAlign.xlHAlignCenter,
                XlHAlign.xlHAlignCenter,
                XlHAlign.xlHAlignCenter,
                XlHAlign.xlHAlignCenter,
                XlHAlign.xlHAlignLeft
            };

            for (int i = 0; i < columnWidths.Length; i++)
            {
                int col = i + 2;
                cablesSheet.Columns[col].ColumnWidth = columnWidths[i];
                ((Range)cablesSheet.Columns[col]).HorizontalAlignment = columnAlignments[i];
            }

            // Set font
            _helper.ApplyFont(cablesSheet.Range["B6", "Q1000"], "Arial", 10);

            // Header row
            Range headerCell = _helper.MergeAndCenter(cablesSheet, 2, 2, 2, 17);
            headerCell.Value = "PROJECT CABLE SCHEDULE";
            _helper.FillAndStyleHeader(headerCell, "#E4A444", 18);

            // Column headers (gray)
            string gray = "#D9D9D9";

            _helper.MergeAndStyle(cablesSheet, 3, 2, 4, 3, "CABLE ID", gray);
            _helper.MergeAndStyle(cablesSheet, 3, 4, 4, 4, "FROM", gray);
            _helper.MergeAndStyle(cablesSheet, 3, 5, 4, 5, "TO", gray);
            _helper.MergeAndStyle(cablesSheet, 3, 6, 4, 6, "QTY", gray);
            _helper.MergeAndStyle(cablesSheet, 3, 7, 4, 7, "SIZE", gray);
            _helper.MergeAndStyle(cablesSheet, 3, 8, 4, 8, "TYPE", gray);
            _helper.MergeAndStyle(cablesSheet, 3, 9, 4, 9, "GND", gray);
            _helper.MergeAndStyle(cablesSheet, 3, 10, 4, 10, "DESCRIPTION", gray);
            _helper.MergeAndStyle(cablesSheet, 3, 11, 4, 11, "RACEWAY ROUTING", gray);
            _helper.MergeAndStyle(cablesSheet, 3, 12, 4, 12, "SIGNAL TYPE", gray);

            // PLC I/O Header
            _helper.MergeAndStyle(cablesSheet, 3, 13, 3, 16, "PLC I/O COUNTS", gray);
            _helper.StyleSingleCell(cablesSheet, 4, 13, "DI", gray);
            _helper.StyleSingleCell(cablesSheet, 4, 14, "DO", gray);
            _helper.StyleSingleCell(cablesSheet, 4, 15, "AI", gray);
            _helper.StyleSingleCell(cablesSheet, 4, 16, "LOCATION", gray);

            // Comments header
            _helper.MergeAndStyle(cablesSheet, 3, 17, 4, 17, "COMMENTS", gray);

            // Subsection row
            Range subSection = _helper.MergeAndCenter(cablesSheet, 5, 2, 5, 17);
            subSection.Value = "000 SERIES - GENERAL DISTRIBUTION";
            _helper.FillAndStyleHeader(subSection, gray, 10, true, XlHAlign.xlHAlignLeft);

            // Example data row
            Range rowLabel = _helper.MergeAndCenter(cablesSheet, 6, 2, 6, 3);
            rowLabel.Value = "UTILITY POWER";
            rowLabel.Font.Bold = true;

            // Table data border styling
            for (int row = 6; row <= 30; row++)
            {
                for (int col = 2; col <= 17; col++)
                {
                    Range cell = cablesSheet.Cells[row, col];
                    _helper.ApplyThinTableBorder(cell);

                    if (col == 2)
                        cell.Borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlMedium;

                    if (col == 17)
                        cell.Borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlMedium;

                    if (row == 30)
                        cell.Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlMedium;
                }
            }

            Marshal.ReleaseComObject(cablesSheet);
        }
    }

}
