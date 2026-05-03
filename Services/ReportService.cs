using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using StripTestBlazor.Models;

namespace StripTestBlazor.Services
{
    public enum FileStatus { Waiting, Processing, Done, Error }

    public class FileEntry
    {
        public string     Name    { get; set; } = "";
        public byte[]     Data    { get; set; } = Array.Empty<byte>();
        public FileStatus Status  { get; set; } = FileStatus.Waiting;
        public LogResult? Result  { get; set; }
    }

    public class ReportService
    {
        public List<FileEntry>  Files      { get; } = new();
        public bool             IsRunning  { get; private set; }
        public bool             IsPaused   { get; private set; }
        public int              Progress   { get; private set; }

        public event Action?    StateChanged;

        TaskCompletionSource<bool> _pauseTcs =
            new(TaskCreationOptions.RunContinuationsAsynchronously);
        CancellationTokenSource _cts = new();

        // ── File management ───────────────────────────────────────
        public void AddFiles(IEnumerable<(string name, byte[] data)> newFiles)
        {
            foreach (var (name, data) in newFiles)
            {
                string ext = Path.GetExtension(name).ToLowerInvariant();
                if ((ext != ".xls" && ext != ".xlsx") || name.StartsWith("~$")) continue;
                if (Files.Any(f => f.Name == name)) continue;
                Files.Add(new FileEntry { Name = name, Data = data });
            }
            Notify();
        }

        public void ClearFiles()
        {
            if (IsRunning) return;
            Files.Clear();
            Progress = 0;
            Notify();
        }

        // ── Processing ────────────────────────────────────────────
        public async Task RunAsync()
        {
            if (IsRunning || Files.Count == 0) return;

            IsRunning = true;
            IsPaused  = false;
            Progress  = 0;
            _cts      = new CancellationTokenSource();
            _pauseTcs = new TaskCompletionSource<bool>(
                TaskCreationOptions.RunContinuationsAsynchronously);
            _pauseTcs.SetResult(true);

            // Reset all to waiting
            foreach (var f in Files) { f.Status = FileStatus.Waiting; f.Result = null; }
            Notify();

            for (int i = 0; i < Files.Count; i++)
            {
                if (_cts.Token.IsCancellationRequested) break;

                // Async pause
                await _pauseTcs.Task.ConfigureAwait(false);
                if (_cts.Token.IsCancellationRequested) break;

                var entry = Files[i];
                entry.Status = FileStatus.Processing;
                Progress = i;
                Notify();

                // Parse off the synchronous path (yields to UI between files)
                await Task.Yield();

                try
                {
                    entry.Result = LogParser.Parse(entry.Data, entry.Name);
                    entry.Status = entry.Result.Error == null
                        ? FileStatus.Done : FileStatus.Error;
                }
                catch (Exception ex)
                {
                    entry.Result = new LogResult { FileName = entry.Name, Error = ex.Message };
                    entry.Status = FileStatus.Error;
                }

                Progress = i + 1;
                Notify();
            }

            IsRunning = false;
            IsPaused  = false;
            Notify();
        }

        public void Pause()
        {
            if (!IsRunning || IsPaused) return;
            IsPaused  = true;
            _pauseTcs = new TaskCompletionSource<bool>(
                TaskCreationOptions.RunContinuationsAsynchronously);
            Notify();
        }

        public void Resume()
        {
            if (!IsPaused) return;
            IsPaused = false;
            _pauseTcs.TrySetResult(true);
            Notify();
        }

        public void Stop()
        {
            _cts.Cancel();
            _pauseTcs.TrySetResult(true);
            IsRunning = false;
            IsPaused  = false;
            Notify();
        }

        // ── Summary helpers ───────────────────────────────────────
        public IEnumerable<LogResult> GoodResults =>
            Files.Where(f => f.Result?.Error == null && f.Result != null)
                 .Select(f => f.Result!);

        public int TotalInput  => GoodResults.Sum(r => r.Input);
        public int TotalOutput => GoodResults.Sum(r => r.Output);
        public double TotalYield => TotalInput > 0 ? (double)TotalOutput / TotalInput : 0;

        // ── Excel generation ──────────────────────────────────────
        public byte[] GenerateReport()
        {
            var results = GoodResults.ToList();
            var reportDate = DateTime.Today;

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var pkg = new ExcelPackage();
            var ws = pkg.Workbook.Worksheets.Add("工作表1");

            // Colors
            var hdrBg  = ColorTranslator.FromHtml("#4472C4");
            var hdrFg  = Color.White;
            var rowOdd = ColorTranslator.FromHtml("#DCE6F1");
            var border = ColorTranslator.FromHtml("#95B3D7");
            var warn   = ColorTranslator.FromHtml("#FFFF00");
            var alert  = ColorTranslator.FromHtml("#FF0000");

            int col = 1;
            void AddSingle(string label)
            {
                ws.Cells[1, col].Value = label;
                StyleHdr(ws.Cells[1, col], hdrBg, hdrFg, border);
                col++;
            }
            void AddMerged(string label)
            {
                ws.Cells[1, col, 1, col + 1].Merge = true;
                ws.Cells[1, col].Value = label;
                StyleHdr(ws.Cells[1, col], hdrBg, hdrFg, border);
                col += 2;
            }

            AddSingle("Test Date/ Time"); AddSingle("Report Date"); AddSingle("Lot No.");
            AddSingle("Input"); AddSingle("Output"); AddSingle("Yield");
            AddMerged("VTH"); AddMerged("VTH < 1.5 V"); AddMerged("VTH > 6 V");
            AddMerged("VFSD"); AddMerged("BVDSS 2"); AddMerged("BVDSS 1");
            AddMerged("Delta3"); AddMerged("IPD-10V");
            AddMerged("IDSS1"); AddMerged("IDSS2"); AddMerged("IDSS3");

            int row = 2;
            foreach (var r in results)
            {
                bool odd = (row % 2 == 0);
                int c = 1;

                void Str(string s)   { ws.Cells[row,c].Value=s; StyleDat(ws.Cells[row,c],odd,rowOdd,border,null); c++; }
                void Dt(DateTime dt) { ws.Cells[row,c].Value=dt.Date; ws.Cells[row,c].Style.Numberformat.Format="yyyy/MM/dd"; StyleDat(ws.Cells[row,c],odd,rowOdd,border,null); c++; }
                void Int(int v)      { ws.Cells[row,c].Value=v; StyleDat(ws.Cells[row,c],odd,rowOdd,border,null); c++; }
                void Yld(double v)   { ws.Cells[row,c].Value=v; ws.Cells[row,c].Style.Numberformat.Format="0.00%"; StyleDat(ws.Cells[row,c],odd,rowOdd,border,null); c++; }
                void Pair(int fail, int total)
                {
                    double rate = total > 0 ? (double)fail / total : 0;
                    Color? hi = rate > 0.01 ? alert : rate > 0.005 ? warn : (Color?)null;
                    ws.Cells[row,c].Value=fail; StyleDat(ws.Cells[row,c],odd,rowOdd,border,hi); c++;
                    ws.Cells[row,c].Value=rate; ws.Cells[row,c].Style.Numberformat.Format="0.00%"; StyleDat(ws.Cells[row,c],odd,rowOdd,border,hi); c++;
                }

                Str(r.TestDateTime); Dt(reportDate); Str(r.LotNo);
                Int(r.Input); Int(r.Output); Yld(r.Yield);
                Pair(r.FailCounts.GetValueOrDefault("VTH"),     r.Input);
                Pair(r.VthLt15,                                  r.Input);
                Pair(r.VthGt6,                                   r.Input);
                Pair(r.FailCounts.GetValueOrDefault("VFSD"),    r.Input);
                Pair(r.FailCounts.GetValueOrDefault("BVDSS 2"), r.Input);
                Pair(r.FailCounts.GetValueOrDefault("BVDSS 1"), r.Input);
                Pair(r.FailCounts.GetValueOrDefault("Delta3"),  r.Input);
                Pair(r.FailCounts.GetValueOrDefault("IPD-10V"), r.Input);
                Pair(r.FailCounts.GetValueOrDefault("IDSS1"),   r.Input);
                Pair(r.FailCounts.GetValueOrDefault("IDSS2"),   r.Input);
                Pair(r.FailCounts.GetValueOrDefault("IDSS3"),   r.Input);
                row++;
            }

            ws.Column(1).Width=18; ws.Column(2).Width=13; ws.Column(3).Width=22;
            ws.Column(4).Width=9;  ws.Column(5).Width=9;  ws.Column(6).Width=9;
            for (int i=7; i<=28; i++) ws.Column(i).Width=9;
            ws.View.FreezePanes(2,1);
            ws.Cells[1,1,1,col-1].AutoFilter=true;

            return pkg.GetAsByteArray();
        }

        static void StyleHdr(ExcelRange cell, Color bg, Color fg, Color border)
        {
            cell.Style.Font.Bold=true;
            cell.Style.Font.Color.SetColor(fg);
            cell.Style.Fill.PatternType=ExcelFillStyle.Solid;
            cell.Style.Fill.BackgroundColor.SetColor(bg);
            cell.Style.HorizontalAlignment=ExcelHorizontalAlignment.Center;
            cell.Style.VerticalAlignment=ExcelVerticalAlignment.Center;
            cell.Style.WrapText=true;
            SetBorder(cell, border);
        }

        static void StyleDat(ExcelRange cell, bool odd, Color rowOdd, Color borderColor, Color? hi)
        {
            Color bg = hi ?? (odd ? rowOdd : Color.White);
            cell.Style.Fill.PatternType=ExcelFillStyle.Solid;
            cell.Style.Fill.BackgroundColor.SetColor(bg);
            cell.Style.Font.Color.SetColor(hi==Color.FromArgb(255,0,0) ? Color.White : Color.Black);
            cell.Style.HorizontalAlignment=ExcelHorizontalAlignment.Center;
            cell.Style.VerticalAlignment=ExcelVerticalAlignment.Center;
            SetBorder(cell, borderColor);
        }

        static void SetBorder(ExcelRange cell, Color color)
        {
            foreach (var b in new[]{ cell.Style.Border.Top, cell.Style.Border.Bottom,
                                     cell.Style.Border.Left, cell.Style.Border.Right })
            { b.Style=ExcelBorderStyle.Thin; b.Color.SetColor(color); }
        }

        void Notify() => StateChanged?.Invoke();
    }
}
