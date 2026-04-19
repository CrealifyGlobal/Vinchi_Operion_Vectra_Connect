using System;
using System.IO;
using ClosedXML.Excel;
using VectraConnect.Models;

namespace VectraConnect
{
    /// <summary>
    /// Writes a ProjectSchema to a styled .xlsx workbook (Tasks / Resources / Assignments sheets)
    /// and optionally a companion CSV per sheet.
    ///
    /// PRIMARY KEY — VectraKey (Column A on every sheet)
    /// Format:  {MppFileName}_{yyyyMMdd}_{HHmmss}_{SheetCode}_{RowUID:D5}
    /// Example: ConstructionPhase2_20240418_143022_TSK_00042
    ///
    /// Sheet codes: TSK = Tasks | RES = Resources | ASN = Assignments
    ///
    /// The original MS Project ID and UniqueID columns sit immediately after
    /// the VectraKey so native identifiers remain visible alongside it.
    /// </summary>
    public static class SchemaExporter
    {
        // ── Colours ───────────────────────────────────────────────────────────
        private static readonly XLColor HeaderBg    = XLColor.FromHtml("#1F3864");
        private static readonly XLColor HeaderFg    = XLColor.White;
        private static readonly XLColor KeyHeaderBg = XLColor.FromHtml("#0A2140");
        private static readonly XLColor KeyBg       = XLColor.FromHtml("#0D1B2A");
        private static readonly XLColor KeyFg       = XLColor.FromHtml("#00E5FF");
        private static readonly XLColor SummaryBg   = XLColor.FromHtml("#D9E1F2");
        private static readonly XLColor MilestoneBg = XLColor.FromHtml("#FFF2CC");
        private static readonly XLColor AltRowBg    = XLColor.FromHtml("#F5F7FA");

        private const string CodeTask       = "TSK";
        private const string CodeResource   = "RES";
        private const string CodeAssignment = "ASN";

        // ── Public entry point ────────────────────────────────────────────────

        public static ExportResult Export(ProjectSchema schema, string outputFolder, bool includeCsv = true)
        {
            Directory.CreateDirectory(outputFolder);

            string safeName  = SanitizeFileName(schema.MppFileName);
            string timestamp = schema.ExportedAt.ToString("yyyyMMdd_HHmmss");
            string xlsxPath  = Path.Combine(outputFolder, $"{safeName}_{timestamp}.xlsx");

            using (var wb = new XLWorkbook())
            {
                wb.Properties.Author  = "VectraConnect";
                wb.Properties.Title   = schema.ProjectName;
                wb.Properties.Created = schema.ExportedAt;

                BuildCoverSheet(wb, schema, safeName, timestamp);
                BuildTasksSheet(wb, schema, safeName, timestamp);
                BuildResourcesSheet(wb, schema, safeName, timestamp);
                BuildAssignmentsSheet(wb, schema, safeName, timestamp);

                wb.SaveAs(xlsxPath);
            }

            var result = new ExportResult { XlsxPath = xlsxPath };

            if (includeCsv)
            {
                result.TasksCsvPath       = WriteCsv(schema, outputFolder, safeName, timestamp, "Tasks");
                result.ResourcesCsvPath   = WriteCsv(schema, outputFolder, safeName, timestamp, "Resources");
                result.AssignmentsCsvPath = WriteCsv(schema, outputFolder, safeName, timestamp, "Assignments");
            }

            return result;
        }

        // ── VectraKey helpers ─────────────────────────────────────────────────

        private static string MakeKey(string safeName, string timestamp, string sheetCode, int rowUid)
            => $"{safeName}_{timestamp}_{sheetCode}_{rowUid:D5}";

        /// Cantor pairing — unique integer for every (taskUid, resourceUid) pair.
        private static int CompositeUid(int taskUid, int resourceUid)
        {
            long k = ((long)(taskUid + resourceUid) * (taskUid + resourceUid + 1) / 2) + resourceUid;
            return (int)(k & 0x7FFFFFFF);
        }

        // ── Cover sheet ───────────────────────────────────────────────────────

        private static void BuildCoverSheet(XLWorkbook wb, ProjectSchema s,
                                            string safeName, string timestamp)
        {
            var ws = wb.Worksheets.Add("Summary");

            ws.Cell("B2").Value = s.ProjectName;
            ws.Cell("B2").Style.Font.Bold      = true;
            ws.Cell("B2").Style.Font.FontSize  = 18;
            ws.Cell("B2").Style.Font.FontColor = XLColor.FromHtml("#1F3864");

            ws.Cell("B3").Value = $"Exported: {s.ExportedAt:dd MMM yyyy HH:mm:ss}";
            ws.Cell("B3").Style.Font.FontColor = XLColor.Gray;
            ws.Cell("B4").Value = $"Source file: {s.FilePath}";
            ws.Cell("B4").Style.Font.FontColor = XLColor.Gray;
            ws.Cell("B4").Style.Font.Italic    = true;
            ws.Cell("B5").Value = $"Session key prefix: {safeName}_{timestamp}";
            ws.Cell("B5").Style.Font.Bold      = true;
            ws.Cell("B5").Style.Font.FontColor = XLColor.FromHtml("#1F3864");

            int row = 8;
            WriteKvp(ws, row++, "Project Start",   s.Summary.StartDate?.ToString("dd MMM yyyy") ?? "—");
            WriteKvp(ws, row++, "Project Finish",  s.Summary.FinishDate?.ToString("dd MMM yyyy") ?? "—");
            WriteKvp(ws, row++, "% Complete",      $"{s.Summary.PercentComplete:F1}%");
            WriteKvp(ws, row++, "Total Tasks",     s.Summary.TotalTasks.ToString());
            WriteKvp(ws, row++, "Total Resources", s.Summary.TotalResources.ToString());
            WriteKvp(ws, row++, "Currency",        s.Summary.Currency);

            ws.Column("B").Width = 28;
            ws.Column("C").Width = 40;
            ws.SheetView.ShowGridLines = false;
        }

        private static void WriteKvp(IXLWorksheet ws, int row, string label, string value)
        {
            ws.Cell(row, 2).Value = label;
            ws.Cell(row, 2).Style.Font.Bold = true;
            ws.Cell(row, 3).Value = value;
        }

        // ── Tasks sheet ───────────────────────────────────────────────────────
        // Col A: VectraKey | B: ID | C: UniqueID | D onwards: data

        private static void BuildTasksSheet(XLWorkbook wb, ProjectSchema s,
                                            string safeName, string timestamp)
        {
            var ws = wb.Worksheets.Add("Tasks");

            string[] headers = {
                "VectraKey",                                              // A
                "ID", "UniqueID",                                         // B–C
                "WBS", "Outline Level", "Name",                           // D–F
                "Summary?", "Milestone?", "Critical?",                    // G–I
                "Duration", "Start", "Finish", "% Complete",              // J–M
                "Baseline Cost", "Actual Cost", "Remaining Cost",         // N–P
                "Total Slack (d)", "Constraint Type", "Constraint Date",  // Q–S
                "Predecessors", "Notes",                                  // T–U
                "Proposed Changes"                                        // V ← reviewers write here
            };

            WriteHeaderRow(ws, 1, headers);

            int row = 2;
            foreach (var t in s.Tasks)
            {
                ws.Cell(row,  1).Value = MakeKey(safeName, timestamp, CodeTask, t.UniqueID);
                ws.Cell(row,  2).Value = t.ID;
                ws.Cell(row,  3).Value = t.UniqueID;
                ws.Cell(row,  4).Value = t.WBS;
                ws.Cell(row,  5).Value = t.OutlineLevel;
                ws.Cell(row,  6).Value = t.Name;
                ws.Cell(row,  7).Value = t.IsSummary   ? "Yes" : "No";
                ws.Cell(row,  8).Value = t.IsMilestone ? "Yes" : "No";
                ws.Cell(row,  9).Value = t.IsCritical  ? "Yes" : "No";
                ws.Cell(row, 10).Value = t.Duration;
                ws.Cell(row, 11).Value = t.Start?.ToString("yyyy-MM-dd") ?? "";
                ws.Cell(row, 12).Value = t.Finish?.ToString("yyyy-MM-dd") ?? "";
                ws.Cell(row, 13).Value = t.PercentComplete / 100.0;
                ws.Cell(row, 14).Value = t.BaselineCost;
                ws.Cell(row, 15).Value = t.ActualCost;
                ws.Cell(row, 16).Value = t.RemainingCost;
                ws.Cell(row, 17).Value = Math.Round(t.TotalSlack / 480.0, 1);
                ws.Cell(row, 18).Value = t.ConstraintType;
                ws.Cell(row, 19).Value = t.ConstraintDate?.ToString("yyyy-MM-dd") ?? "";
                ws.Cell(row, 20).Value = t.PredecessorIDs;
                ws.Cell(row, 21).Value = t.Notes;

                StyleKeyCell(ws.Cell(row, 1));

                if (t.IsSummary)
                {
                    ws.Range(row, 2, row, 21).Style.Fill.BackgroundColor = SummaryBg;
                    ws.Range(row, 2, row, 21).Style.Font.Bold = true;
                }
                else if (t.IsMilestone)
                {
                    ws.Range(row, 2, row, 21).Style.Fill.BackgroundColor = MilestoneBg;
                }
                else if (row % 2 == 0)
                {
                    ws.Range(row, 2, row, 21).Style.Fill.BackgroundColor = AltRowBg;
                }

                ws.Cell(row, 13).Style.NumberFormat.Format = "0%";
                foreach (int col in new[] { 14, 15, 16 })
                    ws.Cell(row, col).Style.NumberFormat.Format = "#,##0.00";

                row++;
            }

            FormatSheet(ws, headers.Length);
        }

        // ── Resources sheet ───────────────────────────────────────────────────
        // Col A: VectraKey | B: ID | C: UniqueID | D onwards: data

        private static void BuildResourcesSheet(XLWorkbook wb, ProjectSchema s,
                                                string safeName, string timestamp)
        {
            var ws = wb.Worksheets.Add("Resources");

            string[] headers = {
                "VectraKey",                                         // A
                "ID", "UniqueID",                                    // B–C
                "Name", "Type", "Email",                             // D–F
                "Max Units (%)", "Standard Rate", "Overtime Rate",   // G–I
                "Baseline Cost", "Actual Cost",                       // J–K
                "Proposed Changes"                             // L ← reviewers write here
            };

            WriteHeaderRow(ws, 1, headers);

            int row = 2;
            foreach (var r in s.Resources)
            {
                ws.Cell(row,  1).Value = MakeKey(safeName, timestamp, CodeResource, r.UniqueID);
                ws.Cell(row,  2).Value = r.ID;
                ws.Cell(row,  3).Value = r.UniqueID;
                ws.Cell(row,  4).Value = r.Name;
                ws.Cell(row,  5).Value = r.Type;
                ws.Cell(row,  6).Value = r.EmailAddress;
                ws.Cell(row,  7).Value = r.MaxUnits / 100.0;
                ws.Cell(row,  8).Value = r.StandardRate;
                ws.Cell(row,  9).Value = r.OvertimeRate;
                ws.Cell(row, 10).Value = r.BaselineCost;
                ws.Cell(row, 11).Value = r.ActualCost;

                StyleKeyCell(ws.Cell(row, 1));

                if (row % 2 == 0)
                    ws.Range(row, 2, row, 11).Style.Fill.BackgroundColor = AltRowBg;

                ws.Cell(row, 7).Style.NumberFormat.Format = "0%";
                foreach (int col in new[] { 8, 9, 10, 11 })
                    ws.Cell(row, col).Style.NumberFormat.Format = "#,##0.00";

                row++;
            }

            FormatSheet(ws, headers.Length);
        }

        // ── Assignments sheet ─────────────────────────────────────────────────
        // Col A: VectraKey | then Task cols | then Resource cols | then data

        private static void BuildAssignmentsSheet(XLWorkbook wb, ProjectSchema s,
                                                  string safeName, string timestamp)
        {
            var ws = wb.Worksheets.Add("Assignments");

            string[] headers = {
                "VectraKey",                                               // A
                "Task ID", "Task Name", "Task VectraKey",                  // B–D
                "Resource ID", "Resource Name", "Resource VectraKey",      // E–G
                "Units (%)", "Work (hrs)", "Actual Work (hrs)",            // H–J
                "Remaining Work (hrs)", "Cost", "Actual Cost",              // K–M
                "Proposed Changes"                                          // N ← reviewers write here
            };

            WriteHeaderRow(ws, 1, headers);

            int row = 2;
            foreach (var a in s.Assignments)
            {
                int uid = CompositeUid(a.TaskUniqueID, a.ResourceUniqueID);

                ws.Cell(row,  1).Value = MakeKey(safeName, timestamp, CodeAssignment, uid);
                ws.Cell(row,  2).Value = a.TaskUniqueID;
                ws.Cell(row,  3).Value = a.TaskName;
                ws.Cell(row,  4).Value = MakeKey(safeName, timestamp, CodeTask,     a.TaskUniqueID);
                ws.Cell(row,  5).Value = a.ResourceUniqueID;
                ws.Cell(row,  6).Value = a.ResourceName;
                ws.Cell(row,  7).Value = MakeKey(safeName, timestamp, CodeResource, a.ResourceUniqueID);
                ws.Cell(row,  8).Value = a.Units / 100.0;
                ws.Cell(row,  9).Value = Math.Round(a.Work, 2);
                ws.Cell(row, 10).Value = Math.Round(a.ActualWork, 2);
                ws.Cell(row, 11).Value = Math.Round(a.RemainingWork, 2);
                ws.Cell(row, 12).Value = a.Cost;
                ws.Cell(row, 13).Value = a.ActualCost;

                StyleKeyCell(ws.Cell(row, 1));
                StyleKeyCell(ws.Cell(row, 4));  // Task FK
                StyleKeyCell(ws.Cell(row, 7));  // Resource FK

                if (row % 2 == 0)
                    ws.Range(row, 2, row, 13).Style.Fill.BackgroundColor = AltRowBg;

                ws.Cell(row, 8).Style.NumberFormat.Format = "0%";
                foreach (int col in new[] { 12, 13 })
                    ws.Cell(row, col).Style.NumberFormat.Format = "#,##0.00";

                row++;
            }

            FormatSheet(ws, headers.Length);
        }

        // ── Shared helpers ────────────────────────────────────────────────────

        // Amber styling for the Proposed Changes column
        private static readonly XLColor ProposedHeaderBg = XLColor.FromHtml("#BF8F00");
        private static readonly XLColor ProposedHeaderFg = XLColor.White;
        private static readonly XLColor ProposedCellBg   = XLColor.FromHtml("#FFF2CC");

        private static void WriteHeaderRow(IXLWorksheet ws, int row, string[] headers)
        {
            for (int i = 0; i < headers.Length; i++)
            {
                var  cell       = ws.Cell(row, i + 1);
                bool isKey      = headers[i].Contains("VectraKey");
                bool isProposed = headers[i] == "Proposed Changes";

                cell.Value = headers[i];
                cell.Style.Font.Bold            = true;
                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                if (isKey)
                {
                    cell.Style.Font.FontColor       = KeyFg;
                    cell.Style.Fill.BackgroundColor = KeyHeaderBg;
                    cell.Style.Font.FontName        = "Consolas";
                }
                else if (isProposed)
                {
                    cell.Style.Font.FontColor       = ProposedHeaderFg;
                    cell.Style.Fill.BackgroundColor = ProposedHeaderBg;
                    cell.Style.Font.Italic          = true;
                }
                else
                {
                    cell.Style.Font.FontColor       = HeaderFg;
                    cell.Style.Fill.BackgroundColor = HeaderBg;
                }
            }

            // Shade the Proposed Changes data cells amber so the column
            // is immediately visible even when rows are empty
            int proposedColIndex = Array.IndexOf(headers, "Proposed Changes") + 1;
            if (proposedColIndex > 0)
                ws.Column(proposedColIndex).Style.Fill.BackgroundColor = ProposedCellBg;

            ws.Row(row).Height = 18;
        }

        private static void StyleKeyCell(IXLCell cell)
        {
            cell.Style.Fill.BackgroundColor = KeyBg;
            cell.Style.Font.FontColor       = KeyFg;
            cell.Style.Font.FontName        = "Consolas";
            cell.Style.Font.Bold            = false;
            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
        }

        private static void FormatSheet(IXLWorksheet ws, int columnCount)
        {
            ws.SheetView.FreezeRows(1);
            ws.SheetView.FreezeColumns(1);    // VectraKey stays visible while scrolling right
            ws.Columns(1, columnCount).AdjustToContents();

            var range = ws.RangeUsed();
            if (range != null)
            {
                var table = range.CreateTable();
                table.Theme = XLTableTheme.None;
                table.ShowAutoFilter = true;
            }

            ws.SheetView.ShowGridLines = true;
        }

        // ── CSV writer ────────────────────────────────────────────────────────

        private static string WriteCsv(ProjectSchema schema, string folder,
                                       string safeName, string timestamp, string sheetName)
        {
            string path = Path.Combine(folder, $"{safeName}_{sheetName}_{timestamp}.csv");

            using (var sw = new StreamWriter(path, false, System.Text.Encoding.UTF8))
            {
                if (sheetName == "Tasks")
                {
                    sw.WriteLine("VectraKey,ID,UniqueID,WBS,OutlineLevel,Name,Summary,Milestone," +
                                 "Critical,Duration,Start,Finish,PctComplete,BaselineCost,ActualCost," +
                                 "RemainingCost,TotalSlackDays,ConstraintType,ConstraintDate," +
                                 "Predecessors,Notes");

                    foreach (var t in schema.Tasks)
                        sw.WriteLine(Csv(
                            MakeKey(safeName, timestamp, CodeTask, t.UniqueID),
                            t.ID, t.UniqueID, t.WBS, t.OutlineLevel, t.Name,
                            t.IsSummary, t.IsMilestone, t.IsCritical,
                            t.Duration,
                            t.Start?.ToString("yyyy-MM-dd"),
                            t.Finish?.ToString("yyyy-MM-dd"),
                            t.PercentComplete,
                            t.BaselineCost, t.ActualCost, t.RemainingCost,
                            Math.Round(t.TotalSlack / 480.0, 1),
                            t.ConstraintType,
                            t.ConstraintDate?.ToString("yyyy-MM-dd"),
                            t.PredecessorIDs, t.Notes));
                }
                else if (sheetName == "Resources")
                {
                    sw.WriteLine("VectraKey,ID,UniqueID,Name,Type,Email,MaxUnits," +
                                 "StandardRate,OvertimeRate,BaselineCost,ActualCost");

                    foreach (var r in schema.Resources)
                        sw.WriteLine(Csv(
                            MakeKey(safeName, timestamp, CodeResource, r.UniqueID),
                            r.ID, r.UniqueID, r.Name, r.Type, r.EmailAddress,
                            r.MaxUnits, r.StandardRate, r.OvertimeRate,
                            r.BaselineCost, r.ActualCost));
                }
                else if (sheetName == "Assignments")
                {
                    sw.WriteLine("VectraKey,TaskID,TaskName,TaskVectraKey," +
                                 "ResourceID,ResourceName,ResourceVectraKey," +
                                 "Units,WorkHrs,ActualWorkHrs,RemainingWorkHrs,Cost,ActualCost");

                    foreach (var a in schema.Assignments)
                    {
                        int uid = CompositeUid(a.TaskUniqueID, a.ResourceUniqueID);
                        sw.WriteLine(Csv(
                            MakeKey(safeName, timestamp, CodeAssignment, uid),
                            a.TaskUniqueID, a.TaskName,
                            MakeKey(safeName, timestamp, CodeTask,     a.TaskUniqueID),
                            a.ResourceUniqueID, a.ResourceName,
                            MakeKey(safeName, timestamp, CodeResource, a.ResourceUniqueID),
                            a.Units,
                            Math.Round(a.Work, 2),
                            Math.Round(a.ActualWork, 2),
                            Math.Round(a.RemainingWork, 2),
                            a.Cost, a.ActualCost));
                    }
                }
            }

            return path;
        }

        private static string Csv(params object[] values)
        {
            return string.Join(",", Array.ConvertAll(values, v =>
            {
                string s = v?.ToString() ?? "";
                return s.Contains(",") || s.Contains("\"") || s.Contains("\n")
                    ? $"\"{s.Replace("\"", "\"\"")}\"" : s;
            }));
        }

        private static string SanitizeFileName(string name)
        {
            foreach (char c in Path.GetInvalidFileNameChars())
                name = name.Replace(c, '_');
            return name.Trim().Replace(' ', '_');
        }
    }

    public class ExportResult
    {
        public string XlsxPath           { get; set; }
        public string TasksCsvPath        { get; set; }
        public string ResourcesCsvPath    { get; set; }
        public string AssignmentsCsvPath  { get; set; }
    }
}
