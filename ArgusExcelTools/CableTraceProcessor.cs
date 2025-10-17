using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace ArgusExcelTools
{
    internal class CableTraceProcessor
    {
        private readonly CableLibrary cableLibrary = new CableLibrary();
        private readonly ConduitLibrary conduitLibrary = new ConduitLibrary();
        private static string _logFilePath = null;

        public TraceResult Run(Excel.Workbook workbook)
        {
            if (workbook == null)
            {
                throw new ArgumentNullException(nameof(workbook));
            }

            var result = new TraceResult();

            var cablesSheet = workbook.Worksheets.Cast<Excel.Worksheet>().FirstOrDefault(w => w.Name == "Cables");
            var racewaySheet = workbook.Worksheets.Cast<Excel.Worksheet>().FirstOrDefault(w => w.Name == "Raceway");
            var ductbankSheet = workbook.Worksheets.Cast<Excel.Worksheet>().FirstOrDefault(w => w.Name == "Ductbank");
            var traySheet = workbook.Worksheets.Cast<Excel.Worksheet>().FirstOrDefault(w => w.Name == "Cable Tray");
            if (cablesSheet == null || racewaySheet == null)
            {
                throw new InvalidOperationException("Required worksheets 'Cables' and 'Raceway' are missing.");
            }

            var ductbanks = ductbankSheet != null ? BuildDuctbankList(ductbankSheet) : new List<Ductbank>();
            var trays = traySheet != null ? BuildCableTrayList(traySheet) : new List<CableTray>();

            BuildCableList(cablesSheet, result.Cables);
            BuildRacewayList(racewaySheet, result.Raceways);
            ApplyRacewaySizing(racewaySheet, result);

            if (string.IsNullOrEmpty(workbook.Path))
            {
                workbook.Application.Dialogs[Microsoft.Office.Interop.Excel.XlBuiltInDialog.xlDialogSaveAs].Show();
            }
            var folder = workbook.Path;
            if (!string.IsNullOrEmpty(folder))
            {
                ValidateCableTray(trays, result.Cables, folder);
                ValidateDuctbank(ductbanks, result.Raceways, folder);
                ValidateCableRaceway(result.Cables, result.Raceways, folder);
            }

            return result;
        }

        private void BuildCableList(Excel.Worksheet sheet, List<Cable> cables)
        {
            // Check that all expected headers exist
            string[] requiredHeaders = { "CABLE ID", "FROM", "TO", "QTY", "SIZE", "TYPE", "GND", "RACEWAY ROUTING", "SIGNAL TYPE" };
            var headerMap = new Dictionary<string, int>();

            foreach (string header in requiredHeaders)
            {
                int colIndex = FindColumnByHeader(sheet, header);
                headerMap[header] = colIndex;

                if (colIndex == -1)
                {
                    throw new InvalidOperationException(
                        $"'{header}' column was not identified. Expected column '{header}' to be present in the worksheet."
                    );
                }
            }

            int idCol = headerMap["CABLE ID"] + 1;
            int fromCol = headerMap["FROM"];
            int toCol = headerMap["TO"];
            int qtyCol = headerMap["QTY"];
            int sizeCol = headerMap["SIZE"];
            int typeCol = headerMap["TYPE"];
            int groundCol = headerMap["GND"];
            int routingCol = headerMap["RACEWAY ROUTING"];
            int signalCol = headerMap["SIGNAL TYPE"];

            var regex = new Regex(@"^C-\d{1,4}$", RegexOptions.IgnoreCase);
            int lastRow = sheet.UsedRange.Rows.Count;

            for (int row = 2; row <= lastRow; row++)
            {
                var id = ((Excel.Range)sheet.Cells[row, idCol]).Text as string;
                if (!regex.IsMatch(id ?? string.Empty))
                    continue;

                cables.Add(new Cable
                {
                    ID = id,
                    From = ((Excel.Range)sheet.Cells[row, fromCol]).Text as string,
                    To = ((Excel.Range)sheet.Cells[row, toCol]).Text as string,
                    Quantity = ((Excel.Range)sheet.Cells[row, qtyCol]).Text as string,
                    Size = ((Excel.Range)sheet.Cells[row, sizeCol]).Text as string,
                    Type = ((Excel.Range)sheet.Cells[row, typeCol]).Text as string,
                    Ground = ((Excel.Range)sheet.Cells[row, groundCol]).Text as string,
                    RacewayRouting = ((Excel.Range)sheet.Cells[row, routingCol]).Text as string,
                    SignalType = ((Excel.Range)sheet.Cells[row, signalCol]).Text as string
                });
            }
        }


        private void BuildRacewayList(Excel.Worksheet sheet, List<Raceway> raceways)
        {
            // Check that all expected headers exist
            string[] requiredHeaders = { "RACEWAY ID", "RACEWAY SIZE", "FROM", "TO", "CIRCUIT TYPE", "CABLE FILL", "DUCTBANK ROUTING" };
            var headerMap = new Dictionary<string, int>();

            foreach (string header in requiredHeaders)
            {
                int colIndex = FindColumnByHeader(sheet, header);
                headerMap[header] = colIndex;

                if (colIndex == -1)
                {
                    throw new InvalidOperationException(
                        $"'{header}' column was not identified. Expected column '{header}' to be present in the worksheet."
                    );
                }
            }

            int idCol = headerMap["RACEWAY ID"] + 1; 
            int sizeCol = headerMap["RACEWAY SIZE"];
            int fromCol = headerMap["FROM"];
            int toCol = headerMap["TO"];
            int circuitCol = headerMap["CIRCUIT TYPE"];
            int fillCol = headerMap["CABLE FILL"];
            int ductbankCol = headerMap["DUCTBANK ROUTING"];

            var regex = new Regex(@"^R-\d{1,4}$", RegexOptions.IgnoreCase);
            int lastRow = sheet.UsedRange.Rows.Count;

            for (int row = 2; row <= lastRow; row++)
            {
                var id = ((Excel.Range)sheet.Cells[row, idCol]).Text as string;
                if (!regex.IsMatch(id ?? string.Empty))
                    continue;

                raceways.Add(new Raceway
                {
                    ID = id,
                    Size = ((Excel.Range)sheet.Cells[row, sizeCol]).Text as string,
                    From = ((Excel.Range)sheet.Cells[row, fromCol]).Text as string,
                    To = ((Excel.Range)sheet.Cells[row, toCol]).Text as string,
                    CircuitType = ((Excel.Range)sheet.Cells[row, circuitCol]).Text as string,
                    CableFill = ((Excel.Range)sheet.Cells[row, fillCol]).Text as string,
                    DuctbankRouting = ((Excel.Range)sheet.Cells[row, ductbankCol]).Text as string
                });
            }
        }


        private List<Ductbank> BuildDuctbankList(Excel.Worksheet sheet)
        {
            var ductbanks = new List<Ductbank>();
            int idCol = FindFirstMatchColumn(sheet, new Regex("^EDB-\\d{1,4}$", RegexOptions.IgnoreCase));
            if (idCol <= 0)
            {
                return ductbanks;
            }

            var racewayRegex = new Regex("^R-\\d{1,4}$", RegexOptions.IgnoreCase);
            int lastRow = sheet.UsedRange.Rows.Count;
            Ductbank current = null;

            for (int row = 1; row <= lastRow; row++)
            {
                var cell = (Excel.Range)sheet.Cells[row, idCol];
                var text = cell.Text as string;
                if (!string.IsNullOrWhiteSpace(text) && new Regex("^EDB-\\d{1,4}$", RegexOptions.IgnoreCase).IsMatch(text))
                {
                    current = new Ductbank { ID = text };
                    ductbanks.Add(current);
                }
                else if ((cell.MergeCells as bool? ?? false) && current != null)
                {
                    // continue rows for the current ductbank
                }
                else
                {
                    current = null;
                }

                if (current != null)
                {
                    int lastCol = sheet.UsedRange.Columns.Count;
                    for (int col = idCol + 1; col <= lastCol; col++)
                    {
                        var raceway = ((Excel.Range)sheet.Cells[row, col]).Text as string;
                        if (racewayRegex.IsMatch(raceway ?? string.Empty))
                        {
                            current.Raceways.Add(raceway);
                        }
                    }
                }
            }
            return ductbanks;
        }

        private List<CableTray> BuildCableTrayList(Excel.Worksheet sheet)
        {
            var trays = new List<CableTray>();
            int idCol = FindFirstMatchColumn(sheet, new Regex("^ECT-\\d{1,4}$", RegexOptions.IgnoreCase));
            if (idCol <= 0)
            {
                return trays;
            }

            var cableRegex = new Regex("^C-\\d{1,4}$", RegexOptions.IgnoreCase);
            int lastRow = sheet.UsedRange.Rows.Count;
            CableTray current = null;

            for (int row = 1; row <= lastRow; row++)
            {
                var cell = (Excel.Range)sheet.Cells[row, idCol];
                var text = cell.Text as string;
                if (!string.IsNullOrWhiteSpace(text) && new Regex("^ECT-\\d{1,4}$", RegexOptions.IgnoreCase).IsMatch(text))
                {
                    current = new CableTray { ID = text };
                    trays.Add(current);
                }
                else if ((cell.MergeCells as bool? ?? false) && current != null)
                {
                    // additional row for same tray
                }
                else
                {
                    current = null;
                }

                if (current != null)
                {
                    int lastCol = sheet.UsedRange.Columns.Count;
                    for (int col = idCol + 1; col <= lastCol; col++)
                    {
                        var cable = ((Excel.Range)sheet.Cells[row, col]).Text as string;
                        if (cableRegex.IsMatch(cable ?? string.Empty))
                        {
                            current.Cables.Add(cable);
                        }
                    }
                }
            }
            return trays;
        }

        private void ApplyRacewaySizing(Excel.Worksheet sheet, TraceResult context)
        {
            int outputCol = FindOrCreateColumn(sheet, "Raceway Sizing");
            var regex = new Regex(@"^R-\d{1,4}$", RegexOptions.IgnoreCase);

            for (int row = 2; row <= sheet.UsedRange.Rows.Count; row++)
            {
                var id = ((Excel.Range)sheet.Cells[row, FindColumnByHeader(sheet, "RACEWAY ID")]).Text as string;
                if (!regex.IsMatch(id ?? string.Empty))
                {
                    continue;
                }

                var raceway = context.Raceways.First(r => r.ID == id);
                var cableIds = raceway.CableFill.Split(new[] { ',', ' ' }, StringSplitOptions.RemoveEmptyEntries);

                double cableFill = 0.0;
                int cableCount = 0;
                foreach (var cableId in cableIds)
                {
                    var cable = context.Cables.FirstOrDefault(c => c.ID.Equals(cableId, StringComparison.OrdinalIgnoreCase));
                    if (cable == null)
                    {
                        continue;
                    }

                    double diameter = GetCableDiameter(cable);
                    if (diameter <= 0)
                    {
                        continue;
                    }

                    if (int.TryParse(cable.Quantity, out int qty))
                    {
                        cableFill += (diameter * diameter * 0.79) * qty;
                        cableCount += qty;
                    }
                }

                var size = conduitFillCalculate(cableFill, cableCount);
                sheet.Cells[row, outputCol] = size;
            }
        }

        private double GetCableDiameter(Cable cable)
        {
            double value;
            if ((cable.Type == "XHHW" || cable.Type == "XHHW-2") && cableLibrary.XHHW.TryGetValue(cable.Size, out value))
            {
                return value;
            }
            if ((cable.Type == "THWN" || cable.Type == "THHN" || cable.Type == "THWN-2") && cableLibrary.THHN.TryGetValue(cable.Size, out value))
            {
                return value;
            }
            if ((cable.Type == "TC-ER" || cable.Type == "TCER" || cable.Type == "OSP") && cableLibrary.TCER.TryGetValue(cable.Size, out value))
            {
                return value;
            }
            if ((cable.Type ?? string.Empty).Contains("BELDEN") || (cable.Size ?? string.Empty).Contains("BELDEN"))
            {
                if (cableLibrary.BELDEN.TryGetValue(cable.Size, out value))
                {
                    return value;
                }
            }
            return 0;
        }

        private string conduitFillCalculate(double cableFill, int cableCount)
        {
            if (cableCount == 0)
            {
                return "Error";
            }

            double limit = cableCount == 1 ? 0.53 : cableCount == 2 ? 0.31 : 0.40;
            foreach (var cur in conduitLibrary.HDPE)
            {
                double percentage = cableFill / cur.Value;
                if (percentage < limit)
                {
                    return cur.Key;
                }
            }
            return "Error";
        }

        private void ValidateCableRaceway(List<Cable> cables, List<Raceway> raceways, string folder)
        {
            var log = new List<string>();
            foreach (var cable in cables)
            {
                bool matchFound = false;
                foreach (var raceway in raceways)
                {
                    if (cable.RacewayRouting != null && cable.RacewayRouting.IndexOf(raceway.ID, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        if (raceway.CableFill == null || raceway.CableFill.IndexOf(cable.ID, StringComparison.OrdinalIgnoreCase) < 0)
                        {
                            log.Add($"Raceway {raceway.ID} is in Cable {cable.ID}, but {cable.ID} is not called out in the raceway tab.");
                        }
                        else
                        {
                            matchFound = true;
                        }
                    }
                }

                if (!matchFound && !cable.RacewayRouting.Contains("ECT") && !cable.RacewayRouting.Equals("-"))
                {
                    log.Add($"Cable {cable.ID} has no corresponding raceway routing.");
                }
            }

            WriteLog("RacewayCableErrorLog.txt", log);
        }

        private void ValidateDuctbank(List<Ductbank> ductbanks, List<Raceway> raceways, string folder)
        {
            var log = new List<string>();
            foreach (var ductbank in ductbanks)
            {
                foreach (var racewayId in ductbank.Raceways)
                {
                    var raceway = raceways.FirstOrDefault(r => r.ID.Equals(racewayId, StringComparison.OrdinalIgnoreCase));
                    if (raceway != null)
                    {
                        if (string.IsNullOrEmpty(raceway.DuctbankRouting) || raceway.DuctbankRouting.IndexOf(ductbank.ID, StringComparison.OrdinalIgnoreCase) < 0)
                        {
                            log.Add($"Raceway {racewayId} is listed in Ductbank {ductbank.ID}, but {ductbank.ID} is not in its DuctbankRouting.");
                        }
                    }
                    else
                    {
                        log.Add($"Raceway {racewayId} is listed in Ductbank {ductbank.ID}, but does not exist in the raceway schedule.");
                    }
                }
            }

            WriteLog("DuctbankErrorLog.txt", log);
        }

        private void ValidateCableTray(List<CableTray> trays, List<Cable> cables, string folder)
        {
            var log = new List<string>();
            foreach (var tray in trays)
            {
                foreach (var cableId in tray.Cables)
                {
                    var cable = cables.FirstOrDefault(c => c.ID.Equals(cableId, StringComparison.OrdinalIgnoreCase));
                    if (cable != null)
                    {
                        if (string.IsNullOrEmpty(cable.RacewayRouting) || cable.RacewayRouting.IndexOf(tray.ID, StringComparison.OrdinalIgnoreCase) < 0)
                        {
                            log.Add($"Cable {cableId} is listed in Cable Tray {tray.ID}, but {tray.ID} is not in its Raceway Routing.");
                        }
                    }
                    else
                    {
                        log.Add($"Cable {cableId} is listed in Cable Tray {tray.ID}, but does not exist in the cable schedule.");
                    }
                }
            }

            WriteLog("CableTrayErrorLog.txt", log);
        }

        private static void WriteLog(string filename, List<string> entries)
        {
            if (string.IsNullOrEmpty(_logFilePath))
            {
                using (var dialog = new FolderBrowserDialog())
                {
                    dialog.Description = "Select folder to save Error Logs";

                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        _logFilePath = dialog.SelectedPath;
                    }
                    else
                    {
                        MessageBox.Show("Log save canceled. No log file created.");
                        return;
                    }
                }
            }
            using (var writer = new StreamWriter(Path.Combine(_logFilePath, filename), false))
            {
                if (entries.Any())
                {
                    foreach (var entry in entries)
                    {
                        writer.WriteLine(entry);
                    }
                }
                else
                {
                    writer.WriteLine("No errors found.");
                }
            }
        }


        private static int FindColumnByHeader(Excel.Worksheet sheet, string header)
        {
            int maxCol = sheet.UsedRange.Columns.Count;
            int maxRow = Math.Min(10, sheet.UsedRange.Rows.Count);
            for (int col = 1; col <= maxCol; col++)
            {
                for (int row = 1; row <= maxRow; row++)
                {
                    var text = ((Excel.Range)sheet.Cells[row, col]).Text as string;
                    if (string.Equals(text?.Trim(), header, StringComparison.OrdinalIgnoreCase))
                    {
                        return col;
                    }
                }
            }
            return -1;
        }

        private static int FindFirstMatchColumn(Excel.Worksheet sheet, Regex pattern)
        {
            int maxCol = sheet.UsedRange.Columns.Count;
            int maxRow = Math.Min(10, sheet.UsedRange.Rows.Count);
            for (int col = 1; col <= maxCol; col++)
            {
                for (int row = 1; row <= maxRow; row++)
                {
                    var text = ((Excel.Range)sheet.Cells[row, col]).Text as string;
                    if (pattern.IsMatch(text ?? string.Empty))
                    {
                        return col;
                    }
                }
            }
            return -1;
        }

        private static int FindOrCreateColumn(Excel.Worksheet sheet, string header)
        {
            int col = FindColumnByHeader(sheet, header);
            if (col > 0)
            {
                return col;
            }
            col = sheet.UsedRange.Columns.Count + 1;
            sheet.Cells[1, col] = header;
            ((Excel.Range)sheet.Cells[1, col]).Font.Bold = true;
            return col;
        }
    }
}
