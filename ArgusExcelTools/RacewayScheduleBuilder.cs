using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ArgusExcelTools
{
    internal class RacewayScheduleBuilder
    {
        private readonly Application _excelApp;
        private readonly Workbook _workbook;
        private readonly CreateScheduleHelper _helper;

        public RacewayScheduleBuilder(Application excelApp, Workbook workbook)
        {
            _excelApp = excelApp;
            _workbook = workbook;
            _helper = new CreateScheduleHelper(_excelApp);
        }

        public void Build()
        {
            Worksheet racewaySheet = _workbook.Sheets.Add();
            racewaySheet.Name = "Raceway";

            // Set column widths and alignment
            double[] columnWidths = { 3, 11, 16, 34, 34, 19, 54, 67, 54 };
            XlHAlign[] columnAlignments = {
                XlHAlign.xlHAlignLeft,
                XlHAlign.xlHAlignLeft,
                XlHAlign.xlHAlignCenter,
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
                racewaySheet.Columns[col].ColumnWidth = columnWidths[i];
                ((Range)racewaySheet.Columns[col]).HorizontalAlignment = columnAlignments[i];
            }

            // Set font
            _helper.ApplyFont(racewaySheet.Range["B6", "Q1000"], "Arial", 10);

            // Header row
            Range headerCell = _helper.MergeAndCenter(racewaySheet, 2, 2, 2, 10);
            headerCell.Value = "PROJECT RACEWAY SCHEDULE";
            _helper.FillAndStyleHeader(headerCell, "#E4A444", 18);

            // Column headers (gray)
            string gray = "#D9D9D9";

            _helper.MergeAndStyle(racewaySheet, 3, 2, 4, 3, "RACEWAY ID", gray);
            _helper.MergeAndStyle(racewaySheet, 3, 4, 4, 4, "RACEWAY SIZE", gray);
            _helper.MergeAndStyle(racewaySheet, 3, 5, 4, 5, "FROM", gray);
            _helper.MergeAndStyle(racewaySheet, 3, 6, 4, 6, "TO", gray);
            _helper.MergeAndStyle(racewaySheet, 3, 7, 4, 7, "CIRCUIT TYPE", gray);
            _helper.MergeAndStyle(racewaySheet, 3, 8, 4, 8, "CABLE FILL", gray);
            _helper.MergeAndStyle(racewaySheet, 3, 9, 4, 9, "DESCRIPTION", gray);
            _helper.MergeAndStyle(racewaySheet, 3, 10, 4, 10, "DUCTBANK ROUTING", gray);
            

            // Subsection row
            Range subSection = _helper.MergeAndCenter(racewaySheet, 5, 2, 5, 10);
            subSection.Value = "000 SERIES - GENERAL DISTRIBUTION";
            _helper.FillAndStyleHeader(subSection, gray, 10, true, XlHAlign.xlHAlignLeft);

            // Example data row
            Range rowLabel = _helper.MergeAndCenter(racewaySheet, 6, 2, 6, 3);
            rowLabel.Value = "UTILITY POWER";
            rowLabel.Font.Bold = true;

            // Table data border styling
            for (int row = 6; row <= 30; row++)
            {
                for (int col = 2; col <= 10; col++)
                {
                    Range cell = racewaySheet.Cells[row, col];
                    _helper.ApplyThinTableBorder(cell);

                    if (col == 2)
                        cell.Borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlMedium;

                    if (col == 10)
                        cell.Borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlMedium;

                    if (row == 30)
                        cell.Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlMedium;
                }
            }

            Marshal.ReleaseComObject(racewaySheet);
        }
    }
}
