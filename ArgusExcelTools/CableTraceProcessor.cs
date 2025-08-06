using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace ArgusElectrical
{
    internal class CableTraceProcessor
    {
        private readonly CableLibrary cableLibrary = new();
        private readonly ConduitLibrary conduitLibrary = new();

        public TraceResult Run(Excel.Workbook workbook)
        {
            if (workbook == null)
            {
                throw new ArgumentNullException(nameof(workbook));
            }

            var result = new TraceResult();

            var cablesSheet = workbook.Worksheets.Cast<Excel.Worksheet>().FirstOrDefault(w => w.Name == "Cables");
            var racewaySheet = workbook.Worksheets.Cast<Excel.Worksheet>().FirstOrDefault(w => w.Name == "Raceway");
            if (cablesSheet == null || racewaySheet == null)
            {
                throw new InvalidOperationException("Required worksheets 'Cables' and 'Raceway' are missing.");
            }

            BuildCableList(cablesSheet, result.Cables);
            BuildRacewayList(racewaySheet, result.Raceways);
            ApplyRacewaySizing(racewaySheet, result);

            return result;
        }

        private void BuildCableList(Excel.Worksheet sheet, List<Cable> cables)
        {
            int idCol = FindColumnByHeader(sheet, "CABLE ID");
            int fromCol = FindColumnByHeader(sheet, "FROM");
            int toCol = FindColumnByHeader(sheet, "TO");
            int qtyCol = FindColumnByHeader(sheet, "QTY");
            int sizeCol = FindColumnByHeader(sheet, "SIZE");
            int typeCol = FindColumnByHeader(sheet, "TYPE");
            int groundCol = FindColumnByHeader(sheet, "GND");
            int routingCol = FindColumnByHeader(sheet, "RACEWAY ROUTING");
            int signalCol = FindColumnByHeader(sheet, "SIGNAL TYPE");

            var regex = new Regex(@"^C-\d{1,4}$", RegexOptions.IgnoreCase);
            int lastRow = sheet.UsedRange.Rows.Count;

            for (int row = 2; row <= lastRow; row++)
            {
                var id = ((Excel.Range)sheet.Cells[row, idCol]).Text as string;
                if (!regex.IsMatch(id ?? string.Empty))
                {
                    continue;
                }

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
            int idCol = FindColumnByHeader(sheet, "RACEWAY ID");
            int sizeCol = FindColumnByHeader(sheet, "RACEWAY SIZE");
            int fromCol = FindColumnByHeader(sheet, "FROM");
            int toCol = FindColumnByHeader(sheet, "TO");
            int circuitCol = FindColumnByHeader(sheet, "CIRCUIT TYPE");
            int fillCol = FindColumnByHeader(sheet, "CABLE FILL");
            int ductbankCol = FindColumnByHeader(sheet, "DUCTBANK ROUTING");

            var regex = new Regex(@"^R-\d{1,4}$", RegexOptions.IgnoreCase);
            int lastRow = sheet.UsedRange.Rows.Count;

            for (int row = 2; row <= lastRow; row++)
            {
                var id = ((Excel.Range)sheet.Cells[row, idCol]).Text as string;
                if (!regex.IsMatch(id ?? string.Empty))
                {
                    continue;
                }

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
