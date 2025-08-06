using System;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace ArgusElectrical
{
    internal class ScheduleGenerator
    {
        public void Generate(Excel.Workbook workbook, TraceResult context)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (context == null) throw new ArgumentNullException(nameof(context));

            var sheet = GetOrCreateSheet(workbook, "Cable Schedule");
            sheet.Cells.Clear();

            string[] headers = { "CABLE ID", "FROM", "TO", "QTY", "SIZE", "TYPE", "GND", "RACEWAY ROUTING", "SIGNAL TYPE" };
            for (int i = 0; i < headers.Length; i++)
            {
                sheet.Cells[1, i + 1] = headers[i];
            }

            int row = 2;
            foreach (var cable in context.Cables)
            {
                sheet.Cells[row, 1] = cable.ID;
                sheet.Cells[row, 2] = cable.From;
                sheet.Cells[row, 3] = cable.To;
                sheet.Cells[row, 4] = cable.Quantity;
                sheet.Cells[row, 5] = cable.Size;
                sheet.Cells[row, 6] = cable.Type;
                sheet.Cells[row, 7] = cable.Ground;
                sheet.Cells[row, 8] = cable.RacewayRouting;
                sheet.Cells[row, 9] = cable.SignalType;
                row++;
            }

            sheet.Columns.AutoFit();
        }

        private Excel.Worksheet GetOrCreateSheet(Excel.Workbook workbook, string name)
        {
            var sheet = workbook.Worksheets.Cast<Excel.Worksheet>().FirstOrDefault(w => w.Name == name);
            if (sheet == null)
            {
                sheet = workbook.Worksheets.Add();
                sheet.Name = name;
            }
            return sheet;
        }
    }
}
