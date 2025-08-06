using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Drawing;
using Excel = Microsoft.Office.Interop.Excel;

namespace ArgusExcelTools
{
    internal class MechCsvProcessor
    {
        private readonly Excel.Application _excelApp;

        public MechCsvProcessor(Excel.Application excelApp)
        {
            _excelApp = excelApp;
        }

        private Excel.Worksheet GetOrCreateSheet(string sheetName)
        {
            Excel.Workbook wb = _excelApp.ActiveWorkbook;

            foreach (Excel.Worksheet sheet in wb.Sheets)
            {
                if (sheet.Name == sheetName)
                {
                    sheet.Cells.Clear();  // Clear all or your range only
                    return sheet;
                }
            }

            Excel.Worksheet newSheet = wb.Sheets.Add();
            newSheet.Name = sheetName;
            return newSheet;
        }

        // Desired headers per type
        private readonly List<string> valveHeaders = new List<string>
        {
            "Count", "Name", "TAG", "SIZE", "DESCRIPTION", "SPEC", "FLOW", "PRESSURE", "TANKTYPE"
        };

        private readonly List<string> vesselHeaders = new List<string>
        {
            "Vessel Tag", "Drawing #", "Vessel Type", "Model Number",
            "Volume (Gal)", "Rate Flow Rate (GPM)", "Pressure Set Point (psi)"
        };

        private readonly List<string> equipmentHeaders = new List<string>
        {
            "Equipment Tag", "Drawing #", "Vessel Type", "Model Number",
            "Inlet Connection Size (inch)", "Inlet Connection Type",
            "Outlet Connection Size (inch)", "Outlet Connection Type",
            "Design Pressure/MA WP (PSIG)", "Design Temperature (F)", "Design Flow Rate (GPM)"
        };

        public void ProcessMechanicalCsv(string csvFilePath)
        {
            var allItems = ParseCsv(csvFilePath);

            var valves = new List<Valve>();
            var vessels = new List<Vessel>();
            var equipment = new List<Equipment>();

            foreach (var row in allItems)
            {
                string name = row.ContainsKey("Name") ? row["Name"].ToLower() : "";

                if (name.Contains("valve"))
                {
                    valves.Add(new Valve { Fields = new Dictionary<string, string>(row) });
                }
                else if (name.Contains("vessel"))
                {
                    vessels.Add(new Vessel { Fields = new Dictionary<string, string>(row) });
                }
                else
                {
                    equipment.Add(new Equipment { Fields = new Dictionary<string, string>(row) });
                }
            }



            WriteValveSchedule(valves);
            WriteVesselSchedule(vessels);
            WriteEquipmentSchedule(equipment);

            MessageBox.Show("Mechanical schedules created successfully!", "Success");
        }

        /// <summary>
        /// Parse CSV rows as generic dictionaries.
        /// </summary>
        private List<Dictionary<string, string>> ParseCsv(string csvFilePath)
        {
            var rows = new List<Dictionary<string, string>>();
            var allHeaders = new List<string>();

            using (var reader = new StreamReader(csvFilePath))
            {
                string[] headerRow = reader.ReadLine().Split(',');

                for (int i = 0; i < headerRow.Length; i++)
                {
                    string current = headerRow[i].Trim();
                    if (string.IsNullOrEmpty(current))
                    {
                        if (i + 1 < headerRow.Length && string.IsNullOrEmpty(headerRow[i + 1]))
                            break; // Stop at two blanks
                    }
                    else
                    {
                        allHeaders.Add(current);
                    }
                }

                while (!reader.EndOfStream)
                {
                    string[] fields = reader.ReadLine().Split(',');
                    var dict = new Dictionary<string, string>();
                    for (int i = 0; i < allHeaders.Count; i++)
                    {
                        string value = (fields.Length > i) ? fields[i].Trim() : "";
                        dict[allHeaders[i]] = value;
                    }
                    rows.Add(dict);
                }
                
            }

            return rows;
        }

        /// <summary>
        /// Build the Valve worksheet.
        /// </summary>
        private void WriteValveSchedule(List<Valve> valves)
        {
            var helper = new CreateScheduleHelper(_excelApp);
            Excel.Worksheet ws = GetOrCreateSheet("Valves");

            helper.FillRange(ws, 1, 1, 1, 10, "#FFFFFF");
            var titleRange = helper.MergeAndCenter(ws, 1, 1, 1, 10);
            titleRange.Value = "VALVE SCHEDULE";
            helper.ApplyFont(titleRange, "Arial", 18);
            titleRange.Font.Bold = true;

            InsertLogo(ws);

            for (int i = 0; i < valveHeaders.Count; i++)
            {
                Excel.Range cell = ws.Cells[2, i + 1];
                cell.Value = valveHeaders[i];
                helper.FillAndStyleHeader(cell, "#757171", 12);
                cell.Font.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#FFC000"));
            }

            int row = 3;
            foreach (var valve in valves)
            {
                string fillHex = (row % 2 == 1) ? "#C9C9C9" : "#DBDBDB";
                for (int col = 0; col < valveHeaders.Count; col++)
                {
                    Excel.Range cell = ws.Cells[row, col + 1];
                    string header = valveHeaders[col];
                    string value = valve.Fields.ContainsKey(header) ? valve.Fields[header] : "";
                    cell.Value = value;

                    cell.Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml(fillHex));
                    cell.Font.Name = "Arial";
                    cell.Font.Size = 10;

                    helper.ApplyThinTableBorder(cell);
                }
                row++;
            }

            ws.Columns.AutoFit();
        }

        /// <summary>
        /// Build the Vessel worksheet.
        /// </summary>
        private void WriteVesselSchedule(List<Vessel> vessels)
        {
            var helper = new CreateScheduleHelper(_excelApp);
            Excel.Worksheet ws = GetOrCreateSheet("Vessels");

            helper.FillRange(ws, 1, 1, 1, 7, "#FFFFFF");
            var titleRange = helper.MergeAndCenter(ws, 1, 1, 1, 7);
            titleRange.Value = "VESSEL SCHEDULE";
            helper.ApplyFont(titleRange, "Arial", 18);
            titleRange.Font.Bold = true;

            InsertLogo(ws);

            for (int i = 0; i < vesselHeaders.Count; i++)
            {
                Excel.Range cell = ws.Cells[2, i + 1];
                cell.Value = vesselHeaders[i];
                helper.FillAndStyleHeader(cell, "#757171", 12);
                cell.Font.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#FFC000"));
            }

            // Mapping for specific output headers
            var headerMap = new Dictionary<string, string>
    {
        { "Vessel Tag", "TAG" },
        { "Drawing #", "DRAWING_NUMBER" }
    };

            int row = 3;
            foreach (var vessel in vessels)
            {
                string fillHex = (row % 2 == 1) ? "#C9C9C9" : "#DBDBDB";

                for (int col = 0; col < vesselHeaders.Count; col++)
                {
                    Excel.Range cell = ws.Cells[row, col + 1];
                    string outputHeader = vesselHeaders[col];

                    // If we have a mapping, use it; else try direct match
                    string sourceHeader = headerMap.ContainsKey(outputHeader) ? headerMap[outputHeader] : outputHeader;

                    string value = vessel.Fields.ContainsKey(sourceHeader) ? vessel.Fields[sourceHeader] : "";
                    cell.Value = value;

                    cell.Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml(fillHex));
                    cell.Font.Name = "Arial";
                    cell.Font.Size = 10;

                    helper.ApplyThinTableBorder(cell);
                }

                row++;
            }

            ws.Columns.AutoFit();
        }

        /// <summary>
        /// Build the Equipment worksheet.
        /// </summary>
        private void WriteEquipmentSchedule(List<Equipment> equipmentList)
        {
            var helper = new CreateScheduleHelper(_excelApp);
            Excel.Worksheet ws = GetOrCreateSheet("Equipment");

            helper.FillRange(ws, 1, 1, 1, 11, "#FFFFFF");
            var titleRange = helper.MergeAndCenter(ws, 1, 1, 1, 11);
            titleRange.Value = "EQUIPMENT SCHEDULE";
            helper.ApplyFont(titleRange, "Arial", 18);
            titleRange.Font.Bold = true;

            InsertLogo(ws);

            for (int i = 0; i < equipmentHeaders.Count; i++)
            {
                Excel.Range cell = ws.Cells[2, i + 1];
                cell.Value = equipmentHeaders[i];
                helper.FillAndStyleHeader(cell, "#757171", 12);
                cell.Font.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#FFC000"));
            }

            int row = 3;
            foreach (var equipment in equipmentList)
            {
                string fillHex = (row % 2 == 1) ? "#C9C9C9" : "#DBDBDB";
                for (int col = 0; col < equipmentHeaders.Count; col++)
                {
                    Excel.Range cell = ws.Cells[row, col + 1];
                    string header = equipmentHeaders[col];
                    string value = equipment.Fields.ContainsKey(header) ? equipment.Fields[header] : "";
                    cell.Value = value;

                    cell.Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml(fillHex));
                    cell.Font.Name = "Arial";
                    cell.Font.Size = 10;

                    helper.ApplyThinTableBorder(cell);
                }
                row++;
            }

            ws.Columns.AutoFit();
        }

        private void InsertLogo(Excel.Worksheet ws)
        {
            string logoPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources", "Argus Logo.png");
            if (File.Exists(logoPath))
            {
                ws.Shapes.AddPicture(
                    logoPath,
                    Microsoft.Office.Core.MsoTriState.msoFalse,
                    Microsoft.Office.Core.MsoTriState.msoCTrue,
                    Left: 10, Top: 5, Width: 80, Height: 40
                );
            }
        }
    }
}