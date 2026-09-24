using Electre_Customize_DotNet.Logs;
using Electre_Customize_DotNet.Objects;
using System.Diagnostics;
using System.Text;

namespace Electre_Customize_DotNet.Helpers.Megger
{
    internal static class MeggerTextReport
    {
        public static bool HighFilesWritten { get; private set; }

        public static void WriteCombinedAndHigh(
            List<ElectreObject> filteredData,
            string reportsFolderPath,
            List<string> lowMeggerData,
            List<(string ConnectorName, string PinNumber)> libraryPins)
        {
            var sw = Stopwatch.StartNew();
            HighFilesWritten = false;

            string templpath = Path.Combine(GlobalVar.StrtCmd, "templ");
            string folderpath = Path.Combine(templpath, reportsFolderPath);

            if (!Directory.Exists(folderpath))
                Directory.CreateDirectory(folderpath);

            if (filteredData == null || filteredData.Count == 0)
                return;

            var seenEnds = new HashSet<string>(filteredData.Count + libraryPins.Count);
            var ends = new List<MeggerEndpoint>(filteredData.Count + libraryPins.Count);

            void TryAddEnd(string conn, string pin, string wire, string subNet, bool hasWire)
            {
                string key = (conn ?? "") + "#" + (pin ?? "");
                if (!seenEnds.Add(key))
                    return;
                ends.Add(new MeggerEndpoint(conn, pin, wire, subNet, hasWire));
            }

            for (int i = 0; i < filteredData.Count; i++)
            {
                var e = filteredData[i];
                TryAddEnd(e.ConnectorName, e.PinNumber, e.WireNumber, e.SubNet, true);
            }
            for (int i = 0; i < libraryPins.Count; i++)
            {
                var p = libraryPins[i];
                TryAddEnd(p.ConnectorName, p.PinNumber, null, null, false);
            }

            int count = ends.Count;
            var lowMeggerKeys = new HashSet<(string, string, string, string)>(lowMeggerData.Count);
            for (int k = 0; k < lowMeggerData.Count; k++)
            {
                string line = lowMeggerData[k];
                int s1 = line.IndexOf(';');
                if (s1 < 0) continue;
                int s2 = line.IndexOf(';', s1 + 1);
                if (s2 < 0) continue;
                int s3 = line.IndexOf(';', s2 + 1);
                if (s3 < 0) continue;
                int s4 = line.IndexOf(';', s3 + 1);
                int end = s4 < 0 ? line.Length : s4;
                lowMeggerKeys.Add((
                    line.Substring(0, s1),
                    line.Substring(s1 + 1, s2 - s1 - 1),
                    line.Substring(s2 + 1, s3 - s2 - 1),
                    line.Substring(s3 + 1, end - s3 - 1)));
            }
            bool hasLowKeys = lowMeggerKeys.Count > 0;

            using (var meggerWriter = new MeggerSplitWriter(folderpath, "Megger"))
            using (var highWriter = new MeggerSplitWriter(folderpath, "HighMegger"))
            {
                for (int k = 0; k < lowMeggerData.Count; k++)
                    meggerWriter.WriteExistingLine(lowMeggerData[k]);

                int dop = Math.Max(1, Environment.ProcessorCount);
                var tempFiles = new string[dop];
                Parallel.For(0, dop, new ParallelOptions { MaxDegreeOfParallelism = dop }, worker =>
                {
                    string tempPath = Path.Combine(folderpath, "_megger_part_" + worker + ".txt");
                    tempFiles[worker] = tempPath;
                    var fs = new FileStream(tempPath, FileMode.Create, FileAccess.Write, FileShare.Read, MeggerSplitWriter.BufferSize, FileOptions.SequentialScan);
                    using (var partWriter = new StreamWriter(fs, new UTF8Encoding(false), MeggerSplitWriter.BufferSize)
                    {
                        AutoFlush = false,
                        NewLine = "\r\n"
                    })
                    {
                        var sb = new StringBuilder(1 << 20);
                        const int flushAt = 1 << 20;
                        for (int i = worker; i < count; i += dop)
                        {
                            MeggerEndpoint e1 = ends[i];
                            for (int j = i + 1; j < count; j++)
                            {
                                MeggerEndpoint n1 = ends[j];
                                if (e1.HasWire && n1.HasWire
                                    && string.Equals(n1.Wire, e1.Wire)
                                    && string.Equals(n1.SubNet, e1.SubNet))
                                    continue;

                                if (hasLowKeys && lowMeggerKeys.Contains((e1.Conn, e1.Pin, n1.Conn, n1.Pin)))
                                    continue;

                                sb.Append(e1.Conn).Append(';')
                                  .Append(e1.Pin).Append(';')
                                  .Append(n1.Conn).Append(';')
                                  .Append(n1.Pin).Append(";High Megger\r\n");
                                if (sb.Length >= flushAt)
                                {
                                    partWriter.Write(sb);
                                    sb.Clear();
                                }
                            }
                        }
                        if (sb.Length > 0)
                            partWriter.Write(sb);
                    }
                });

                for (int t = 0; t < tempFiles.Length; t++)
                {
                    string tempPath = tempFiles[t];
                    if (string.IsNullOrEmpty(tempPath) || !File.Exists(tempPath))
                        continue;
                    meggerWriter.AppendRawFile(tempPath);
                    highWriter.AppendRawFile(tempPath);
                    try { File.Delete(tempPath); } catch { }
                }

                Logging.Info($"MeggerSheet5: ends={count} combined={meggerWriter.TotalLines} high={highWriter.TotalLines} elapsed={sw.ElapsedMilliseconds}ms");
            }

            HighFilesWritten = true;
            Logging.Info("Megger sheet generation completed successfully.");
        }

        public static void WriteTypedSheet(
            string meggerType,
            string reportsFolderPath,
            IEnumerable<string> outputLines)
        {
            if (string.Equals(meggerType, "High", StringComparison.OrdinalIgnoreCase) && HighFilesWritten)
            {
                Logging.Info("High Megger files already streamed by MeggerSheet5; skipping rewrite");
                return;
            }

            string templpath = Path.Combine(GlobalVar.StrtCmd, "templ");
            string folderpath = Path.Combine(templpath, reportsFolderPath);

            if (!Directory.Exists(folderpath))
                Directory.CreateDirectory(folderpath);

            using (var writer = new MeggerSplitWriter(folderpath, meggerType + "Megger"))
            {
                foreach (var line in outputLines)
                    writer.WriteExistingLine(line);
                Logging.Info($"{meggerType} Megger: lines={writer.TotalLines}");
            }

            Logging.Info($"{meggerType} Megger sheet generation completed successfully.");
        }
    }
}
