using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Data;
using System.Linq;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Drawing;
using System.Windows.Forms;

namespace ArgusExcelTools
{
    /// <summary>
    /// Mechanical-only schedule builder.
    /// Reads the ACTIVE worksheet, normalizes TAGs, builds SEARCH, classifies rows
    /// into categories, writes one sheet per category to a NEW workbook, styles tables,
    /// and applies a couple of conditional highlights.
    /// </summary>
    internal class MechScheduleBuilder
    {
        // ---- Appearance constants (Mechanical brand-ish) ----
        private const string HeaderFillHex = "#757171"; // Argus Grey
        private const string HeaderFontHex = "#FFC000"; // Argus Gold
        private const string FontName = "Aptos";
        private const int HeaderFontSize = 11;

        private const string NoTagPlaceholder = "<NO TAG>";

        public void BuildFromActiveSheet()
        {
            var app = Globals.ThisAddIn?.Application;
            if (app == null) { MessageBox.Show("Excel Application not available."); return; }

            var wb = app.ActiveWorkbook;
            var ws = app.ActiveSheet as Excel.Worksheet;

            if (wb == null || ws == null)
            {
                MessageBox.Show("Open a workbook and select a worksheet, then try again.",
                    "Mechanical Schedule", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            bool oldAlerts = app.DisplayAlerts;
            bool oldScreen = app.ScreenUpdating;
            var oldStatus = app.StatusBar;

            try
            {
                app.DisplayAlerts = false;
                app.ScreenUpdating = false;
                app.StatusBar = "Building Mechanical schedules…";

                // 1) Pull active sheet to DataTable
                var dt = WorksheetToDataTable(ws, firstRowHasHeaders: true);

                // 2) Guard rails
                if (!dt.Columns.Contains("Name"))
                    throw new Exception("Expected column 'Name' not found on the active sheet.");

                // 3) Normalize TAG column (backfill) and drop other tag columns
                EnsureTag(dt);

                // 4) SEARCH column (uppercased concat of common columns)
                BuildSearch(dt);

                // 5) Category classification (Mechanical-leaning regex)
                ClassifyMechanical(dt);

                // 6) Sort by Name for readability
                if (dt.Rows.Count > 0)
                    dt = dt.AsEnumerable().OrderBy(r => r["Name"]?.ToString()).CopyToDataTable();

                // 7) Write NEW workbook: one sheet per category + styling
                WriteMechanicalOutputWorkbook(app, dt);

                app.StatusBar = "✅ Mechanical schedules built.";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Mechanical Schedule", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                app.DisplayAlerts = oldAlerts;
                app.ScreenUpdating = oldScreen;
                app.StatusBar = oldStatus;
            }
        }

        // ----------------- Classification -----------------

        private static void ClassifyMechanical(DataTable dt)
        {
            if (!dt.Columns.Contains("Category"))
                dt.Columns.Add("Category", typeof(string));

            foreach (DataRow r in dt.Rows) r["Category"] = "Other";

            // Base categories from your Python + a few mech-flavored terms
            var patterns = new Dictionary<string, Regex>
            {
                ["Valves"] = new Regex(@"\bVALVE(?:S)?\b|\b(?:V|XV|HV|CV|PV|SV|BV|MOV|AOV|EIV)[-\s]?\d+|\bVLV\b",
                                       RegexOptions.IgnoreCase),
                ["Vessels"] = new Regex(@"\b(?:VESSEL|VES|VSL|TANK|SUMP|SURGE)\b|\b(?:T|TK)[-\s]?\d+",
                                        RegexOptions.IgnoreCase),
                ["Pumps"] = new Regex(@"\bPUMP(?:S)?\b|\b(?:P|PU|PMP)[-\s]?\d+",
                                      RegexOptions.IgnoreCase),
                ["Instrumentation"] = new Regex(@"\b(?:INSTRUMENT|CONTROLLER|FUNCTION|LOGIC)\b",
                                                RegexOptions.IgnoreCase),
                ["Equipment"] = new Regex(
                    // original list plus common mech kit
                    @"
                    \b(FOAM|STRAINER|HOSE|SWITCH|MICRONIC|PIG|METER|ORIFICE|DISC|SIGHT|TEST|ACTUATOR|AUTOMATIC|ARRESTOR)\b
                    | \b(?:F|SF|ST)[-\s]?\d+
                    | \b(HEAT\s*EXCHANGER|FILTER|COALESCER|RELIEF\s*VALVE|RUPTURE\s*DISC|PRESSURE\s*SAFETY|AIR\s*RELEASE)\b
                    | \b(SKID|PACKAGE|SEPARATOR|DRYER)\b
                    ",
                    RegexOptions.IgnoreCase | RegexOptions.IgnorePatternWhitespace),
            };

            foreach (DataRow r in dt.Rows)
            {
                var s = r["SEARCH"]?.ToString() ?? "";
                foreach (var kv in patterns)
                {
                    if (kv.Value.IsMatch(s)) { r["Category"] = kv.Key; break; }
                }
            }
        }

        // ----------------- Output workbook -----------------

        private static void WriteMechanicalOutputWorkbook(Excel.Application app, DataTable dt)
        {
            // Category order
            var orderedCats = new List<string> { "Valves", "Vessels", "Pumps", "Instrumentation", "Equipment", "Other" };

            // Per-category column removals (mechanical bias; we KEEP sizing/rating where helpful)
            var drop = new Dictionary<string, string[]>
            {
                ["Valves"] = new[] { "EQUIPMENT", "LOOPNUMBER", "ELECTRICAL", "FLOW", "HEAD", "HP", "REGISTER", "PRESSURE", "VFR", "VOLUME", "BTM", "INFO", "TOP", "VOL", "Category" },
                ["Vessels"] = new[] { "EQUIPMENT", "LOOPNUMBER", "ELECTRICAL", "HEAD", "HP", "REGISTER", "VFR", "BTM", "INFO", "TOP", "Category" }, // keep SIZE/VOLUME/PRESSURE if present
                ["Pumps"] = new[] { "EQUIPMENT", "LOOPNUMBER", "REGISTER", "PRESSURE", "VFR", "VOLUME", "BTM", "INFO", "TOP", "VOL", "SETPOINT", "Category" },
                // For Instrumentation, keep TAG; mechanical often needs it visible
                ["Instrumentation"] = new[] { "ELECTRICAL", "FLOW", "SIZE", "HEAD", "HP", "REGISTER", "PRESSURE", "VFR", "VOLUME", "BTM", "TOP", "VOL", "CONNECTSIZE", "CONNECTRATING", "PCLASS", "Category" },
                // For Equipment, keep SIZE/ratings if present
                ["Equipment"] = new[] { "ELECTRICAL", "FLOW", "HEAD", "HP", "REGISTER", "PRESSURE", "VFR", "VOLUME", "BTM", "TOP", "VOL", "Category" },
            };

            var outWb = app.Workbooks.Add();
            bool wrote = false;

            foreach (var cat in orderedCats)
            {
                var rows = dt.AsEnumerable().Where(r => (r["Category"]?.ToString() ?? "") == cat).ToList();
                if (!rows.Any()) continue;

                var copy = dt.Clone();
                foreach (var r in rows) copy.ImportRow(r);

                // Drop SEARCH + category-specific columns
                if (copy.Columns.Contains("SEARCH")) copy.Columns.Remove("SEARCH");
                if (drop.TryGetValue(cat, out var toDrop))
                    foreach (var c in toDrop.Where(copy.Columns.Contains).ToList()) copy.Columns.Remove(c);
                if (copy.Columns.Count == 0) continue;

                var ws = (Excel.Worksheet)outWb.Worksheets.Add(After: outWb.Worksheets[outWb.Worksheets.Count]);
                ws.Name = SanitizeSheetName(cat);

                WriteDataTable(ws, copy);
                StyleAsTable(ws, cat);
                ApplyConditionalFormats(ws);

                wrote = true;
            }

            if (!wrote)
            {
                var ws = (Excel.Worksheet)outWb.Worksheets[1];
                ws.Name = "Empty";
                ws.Cells[1, 1].Value2 = "No mechanical items found or categorized.";
            }
        }

        // ----------------- Data prep -----------------

        private static void EnsureTag(DataTable dt)
        {
            var tagPriority = new[] { "TAG", "AETAG", "SRTTAG", "CFTAG", "FSTAG", "MFTAG", "PFTAG" };

            if (!dt.Columns.Contains("TAG")) dt.Columns.Add("TAG", typeof(string));

            foreach (DataRow r in dt.Rows)
            {
                string tag = null;
                foreach (var c in tagPriority.Where(dt.Columns.Contains))
                {
                    var v = r[c]?.ToString();
                    if (!string.IsNullOrWhiteSpace(v)) { tag = v; break; }
                }
                r["TAG"] = string.IsNullOrWhiteSpace(tag) ? NoTagPlaceholder : tag;
            }

            foreach (var c in tagPriority.Where(c => c != "TAG" && dt.Columns.Contains(c)).ToList())
                dt.Columns.Remove(c);
        }

        private static void BuildSearch(DataTable dt)
        {
            var searchCols = new[] { "Name", "TAG", "TYPE", "EQUIPMENT", "SERVICE", "DESCRIPTION", "DESC" }
                             .Where(dt.Columns.Contains).ToList();

            if (!dt.Columns.Contains("SEARCH")) dt.Columns.Add("SEARCH", typeof(string));

            foreach (DataRow r in dt.Rows)
            {
                var parts = searchCols.Select(c => r[c]?.ToString() ?? "");
                r["SEARCH"] = string.Join(" ", parts).ToUpperInvariant();
            }
        }

        // ----------------- Excel helpers -----------------

        private static DataTable WorksheetToDataTable(Excel.Worksheet ws, bool firstRowHasHeaders)
        {
            Excel.Range used = ws.UsedRange;
            int rows = used.Rows.Count, cols = used.Columns.Count;

            if (rows == 1 && cols == 1 && used.Value2 == null)
                return new DataTable(); // empty sheet

            object[,] data = (object[,])used.Value2;
            var dt = new DataTable();

            int rowStart = 1;
            if (firstRowHasHeaders)
            {
                for (int c = 1; c <= cols; c++)
                {
                    string h = data[1, c]?.ToString() ?? $"Column{c}";
                    dt.Columns.Add(h);
                }
                rowStart = 2;
            }
            else
            {
                for (int c = 1; c <= cols; c++) dt.Columns.Add($"Column{c}");
            }

            for (int r = rowStart; r <= rows; r++)
            {
                var dr = dt.NewRow();
                for (int c = 1; c <= cols; c++) dr[c - 1] = data[r, c];
                dt.Rows.Add(dr);
            }
            return dt;
        }

        private static void WriteDataTable(Excel.Worksheet ws, DataTable dt)
        {
            // headers (beautified)
            for (int c = 0; c < dt.Columns.Count; c++)
                ws.Cells[1, c + 1].Value2 = BeautifyHeader(dt.Columns[c].ColumnName);

            if (dt.Rows.Count == 0) return;

            // bulk write values
            object[,] arr = new object[dt.Rows.Count, dt.Columns.Count];
            for (int r = 0; r < dt.Rows.Count; r++)
                for (int c = 0; c < dt.Columns.Count; c++)
                    arr[r, c] = dt.Rows[r][c];

            var start = ws.Cells[2, 1];
            var end = ws.Cells[dt.Rows.Count + 1, dt.Columns.Count];
            ws.Range[start, end].Value2 = arr;
        }

        private static string BeautifyHeader(string raw)
        {
            // Specific renames
            var key = Regex.Replace(raw ?? "", "[^A-Za-z0-9]", "").ToUpperInvariant();
            if (key == "PCLASS") return "Class";
            if (key == "CONNECTRATING") return "Connection Rating";
            if (key == "CONNECTSIZE") return "Connection Size";

            var s = (raw ?? "").Replace("_", " ").Trim();
            if (string.IsNullOrEmpty(s)) return "Column";
            return System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(s.ToLower());
        }

        private static void StyleAsTable(Excel.Worksheet ws, string cat)
        {
            Excel.Range used = ws.UsedRange;

            var lo = ws.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, used,
                                        Type.Missing, Excel.XlYesNoGuess.xlYes);
            lo.Name = MakeSafeTableName(cat + "Table");
            lo.TableStyle = "TableStyleMedium15";

            // Header styling
            Excel.Range header = used.Rows[1];
            header.Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml(HeaderFillHex));
            header.Font.Bold = true;
            header.Font.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml(HeaderFontHex));
            header.Font.Name = FontName;
            header.Font.Size = HeaderFontSize;

            // Body font / alignment
            used.Font.Name = FontName;
            used.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            used.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

            used.Columns.EntireColumn.AutoFit();
        }

        private static void ApplyConditionalFormats(Excel.Worksheet ws)
        {
            Excel.Range used = ws.UsedRange;
            int lastRow = used.Row + used.Rows.Count - 1;
            if (lastRow < 2) return;

            // Column A > 1 => red fill
            var colA = ws.Range[ws.Cells[2, 1], ws.Cells[lastRow, 1]];
            var fc1 = colA.FormatConditions.Add(
                Excel.XlFormatConditionType.xlCellValue,
                Excel.XlFormatConditionOperator.xlGreater, "1");
            fc1.Interior.Color = ColorTranslator.ToOle(Color.Red);

            // TAG == "<NO TAG>" => red fill
            int tagCol = FindHeaderColumnIndex(ws, "TAG");
            if (tagCol > 0)
            {
                var tagRange = ws.Range[ws.Cells[2, tagCol], ws.Cells[lastRow, tagCol]];
                var fc2 = (Excel.FormatCondition)tagRange.FormatConditions.Add(
                    Excel.XlFormatConditionType.xlCellValue,
                    Excel.XlFormatConditionOperator.xlEqual,
                    $"=\"{NoTagPlaceholder}\""   // NOTE: needs = and quotes
                );
                fc2.Interior.Color = ColorTranslator.ToOle(Color.Red);
            }
        }

        private static int FindHeaderColumnIndex(Excel.Worksheet ws, string header)
        {
            Excel.Range used = ws.UsedRange;
            int cols = used.Columns.Count;
            for (int c = 1; c <= cols; c++)
            {
                var v = (ws.Cells[1, c].Value2 as object)?.ToString();
                if (v == null) continue;
                var key = Regex.Replace(v, "[^A-Za-z0-9]", "").ToUpperInvariant();
                if (key == header.ToUpperInvariant()) return c;
            }
            return -1;
        }

        private static string SanitizeSheetName(string name)
        {
            var invalid = "[]:*?/\\";
            foreach (var ch in invalid) name = name.Replace(ch.ToString(), "_");
            if (string.IsNullOrWhiteSpace(name)) name = "Sheet";
            return name.Length > 31 ? name.Substring(0, 31) : name;
        }

        private static string MakeSafeTableName(string name)
        {
            var safe = new string(name.Where(char.IsLetterOrDigit).ToArray());
            if (string.IsNullOrEmpty(safe) || !char.IsLetter(safe[0])) safe = "T" + safe;
            return safe.Length > 25 ? safe.Substring(0, 25) : safe;
        }
    }
}
