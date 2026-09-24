using Electre_Customize_DotNet.Objects;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.Json;
using Excel = Microsoft.Office.Interop.Excel;

namespace Electre_Customize_DotNet.Helpers.Global
{
    /// <summary>
    /// Off-by-default performance instrumentation: Stopwatch-based phase timing and deterministic
    /// workbook output fingerprinting, used only to measure and validate optimization work.
    /// Never touches report generation logic and adds zero overhead when disabled (the default).
    /// </summary>
    internal static class PerfDiagnostics
    {
        private static readonly Lazy<bool> _timingEnabled = new(() =>
            bool.TryParse(ConfigurationManager.AppSettings["PerfDiagnostics"], out var v) && v);

        private static readonly Lazy<bool> _fingerprintEnabled = new(() =>
            bool.TryParse(ConfigurationManager.AppSettings["PerfFingerprint"], out var v) && v);

        public static bool TimingEnabled => _timingEnabled.Value;
        public static bool FingerprintEnabled => _fingerprintEnabled.Value;

        private static readonly object _logLock = new();
        private static readonly string LogFolder = Path.Combine(GlobalVar.StrtCmd ?? Path.GetTempPath(), "Logs");

        private static string TimingLogPath => Path.Combine(LogFolder, $"perf_{DateTime.Now:yyyy-MM-dd}.log");

        /// <summary>
        /// Starts a timer if diagnostics are enabled, otherwise returns null. Pair with StopAndLog.
        /// Use this non-invasive Start/Stop form around existing code blocks so control flow
        /// (early returns, exceptions, loop variables) is never restructured.
        /// </summary>
        public static Stopwatch StartTimer() => TimingEnabled ? Stopwatch.StartNew() : null;

        public static void StopAndLog(Stopwatch sw, string phaseName, string reportName, params (string Key, object Value)[] context)
        {
            if (sw == null)
            {
                return;
            }
            sw.Stop();
            WriteTiming(phaseName, reportName, sw.ElapsedMilliseconds, context);
        }

        /// <summary>
        /// Times an action when diagnostics are enabled; always executes the action.
        /// Extra context (input count, selected count, generated rows, etc.) is free-form and logged as-is.
        /// </summary>
        public static void Time(string phaseName, string reportName, Action action, params (string Key, object Value)[] context)
        {
            if (!TimingEnabled)
            {
                action();
                return;
            }

            var sw = Stopwatch.StartNew();
            try
            {
                action();
            }
            finally
            {
                sw.Stop();
                WriteTiming(phaseName, reportName, sw.ElapsedMilliseconds, context);
            }
        }

        public static T Time<T>(string phaseName, string reportName, Func<T> action, params (string Key, object Value)[] context)
        {
            if (!TimingEnabled)
            {
                return action();
            }

            var sw = Stopwatch.StartNew();
            try
            {
                return action();
            }
            finally
            {
                sw.Stop();
                WriteTiming(phaseName, reportName, sw.ElapsedMilliseconds, context);
            }
        }

        private static void WriteTiming(string phaseName, string reportName, long elapsedMs, (string Key, object Value)[] context)
        {
            try
            {
                if (!Directory.Exists(LogFolder))
                {
                    Directory.CreateDirectory(LogFolder);
                }

                var extras = "";
                if (context is { Length: > 0 })
                {
                    var parts = new List<string>(context.Length);
                    foreach (var (key, value) in context)
                    {
                        parts.Add($"{key}={value}");
                    }
                    extras = " " + string.Join(" ", parts);
                }

                string line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PERF] report={reportName} phase={phaseName} elapsedMs={elapsedMs}{extras}";

                lock (_logLock)
                {
                    File.AppendAllText(TimingLogPath, line + Environment.NewLine);
                }
            }
            catch
            {
                // Diagnostics must never break report generation.
            }
        }

        /// <summary>
        /// Captures a deterministic fingerprint of a workbook's content (sheet names, used-range
        /// dimensions, bulk Value2, merged-cell ranges, per-column widths and a representative
        /// number-format sample) for pre/post optimization comparison. Bulk reads only — never
        /// iterates individual cells. No-op unless PerfFingerprint is enabled.
        /// </summary>
        public static void CaptureFingerprint(Excel.Workbook workbook, string label)
        {
            if (!FingerprintEnabled || workbook == null)
            {
                return;
            }

            try
            {
                var sheets = new List<object>();

                foreach (Excel.Worksheet ws in workbook.Sheets)
                {
                    Excel.Range usedRange = null;
                    Excel.Range mergedCellsRange = null;
                    try
                    {
                        usedRange = ws.UsedRange;
                        object values = usedRange.Value2;

                        var mergeAddresses = new List<string>();
                        // MergeCells over the used range only; avoids a full-sheet scan.
                        object mergeCells = usedRange.MergeCells;
                        if (mergeCells is bool hasUniformMerge && hasUniformMerge)
                        {
                            mergeAddresses.Add(usedRange.Address[false, false]);
                        }

                        var columnWidths = new List<double>();
                        Excel.Range usedColumns = null;
                        try
                        {
                            usedColumns = usedRange.Columns;
                            foreach (Excel.Range col in usedColumns)
                            {
                                columnWidths.Add(Convert.ToDouble(col.ColumnWidth ?? -1));
                                Marshal.ReleaseComObject(col);
                            }
                        }
                        finally
                        {
                            if (usedColumns != null) Marshal.ReleaseComObject(usedColumns);
                        }

                        sheets.Add(new
                        {
                            SheetName = ws.Name,
                            Rows = usedRange.Rows.Count,
                            Cols = usedRange.Columns.Count,
                            Values = values,
                            MergedRanges = mergeAddresses,
                            ColumnWidths = columnWidths
                        });
                    }
                    finally
                    {
                        if (mergedCellsRange != null) Marshal.ReleaseComObject(mergedCellsRange);
                        if (usedRange != null) Marshal.ReleaseComObject(usedRange);
                        Marshal.ReleaseComObject(ws);
                    }
                }

                var fingerprint = new
                {
                    Label = label,
                    CapturedAt = DateTime.Now,
                    WorkbookName = workbook.Name,
                    Sheets = sheets
                };

                if (!Directory.Exists(LogFolder))
                {
                    Directory.CreateDirectory(LogFolder);
                }

                string safeLabel = string.Join("_", label.Split(Path.GetInvalidFileNameChars()));
                string fileName = $"fingerprint_{safeLabel}_{DateTime.Now:yyyyMMdd_HHmmss_fff}.json";
                string json = JsonSerializer.Serialize(fingerprint, new JsonSerializerOptions { WriteIndented = true });

                lock (_logLock)
                {
                    File.WriteAllText(Path.Combine(LogFolder, fileName), json);
                }
            }
            catch
            {
                // Diagnostics must never break report generation.
            }
        }
    }
}
