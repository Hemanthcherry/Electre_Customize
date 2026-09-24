#nullable disable
using System;
using System.Collections.Generic;
using System.Configuration;
using Electre_Customize_DotNet.Objects;

namespace Electre_Customize_DotNet.Helpers.Megger
{
    /// <summary>
    /// Application-facing entry point for the Megger text files. modMain keeps the UI concerns
    /// (message boxes, logging, folder resolution) and delegates all data work to this class:
    ///   - WriteCombinedSheets : Megger_{n}.txt      = lowMeggerData lines, then the High Megger pairs
    ///   - WriteTypeSheets     : HighMegger_{n}.txt  = High Megger pairs only (reuses the plan built above)
    ///                           LowMegger_{n}.txt   = lowMeggerData only
    /// </summary>
    internal static class MeggerReportWriter
    {
        // The High Megger pair set built by WriteCombinedSheets and reused by WriteTypeSheets("High").
        // It is a compact plan (distinct connector/pin nodes), not one string per pair.
        private static MeggerHighPairs _highPairs;

        /// <summary>Optional app setting MeggerWriteParallelism; otherwise a small disk-friendly default.</summary>
        public static int WriteParallelism()
        {
            if (int.TryParse(ConfigurationManager.AppSettings["MeggerWriteParallelism"], out int configured) && configured > 0)
            {
                return configured;
            }
            return MeggerFileWriter.DefaultDegreeOfParallelism();
        }

        /// <summary>Forgets the previous run's High Megger plan.</summary>
        public static void Reset() => _highPairs = null;

        public static List<string> WriteCombinedSheets(
            string folder,
            List<ElectreObject> filteredData,
            IReadOnlyList<string> lowLines,
            IReadOnlyList<(string ConnectorName, string PinNumber)> libraryPins,
            Action<long, int> onPlanned = null)
        {
            var rows = new MeggerRow[filteredData.Count];
            for (int i = 0; i < rows.Length; i++)
            {
                var e = filteredData[i];
                rows[i] = new MeggerRow(e.ConnectorName, e.PinNumber, e.WireNumber, e.SubNet);
            }

            var pairs = MeggerHighPairs.Build(rows, lowLines, libraryPins);
            _highPairs = pairs;

            long maxLines = pairs.TotalLines + lowLines.Count;
            long perSheet = MeggerFileWriter.DefaultMaxLinesPerSheet;
            int sheets = (int)Math.Max(1, (maxLines + perSheet - 1) / perSheet);
            onPlanned?.Invoke(maxLines, sheets);

            return MeggerFileWriter.WriteSheets(folder, "Megger_", lowLines, pairs, perSheet, WriteParallelism());
        }

        public static List<string> WriteTypeSheets(string meggerType, string folder, IReadOnlyList<string> lowLines)
        {
            bool high = meggerType == "High";
            return MeggerFileWriter.WriteSheets(
                folder,
                $"{meggerType}Megger_",
                high ? null : lowLines,
                high ? _highPairs : null,
                MeggerFileWriter.DefaultMaxLinesPerSheet,
                WriteParallelism());
        }
    }
}
