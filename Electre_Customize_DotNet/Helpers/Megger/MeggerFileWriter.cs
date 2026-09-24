#nullable disable
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace Electre_Customize_DotNet.Helpers.Megger
{
    internal static class MeggerFileWriter
    {
        public const long DefaultMaxLinesPerSheet = 5_000_000;
        private const int SinkBufferSize = 4 * 1024 * 1024;

        // StreamWriter(path, false, Encoding.UTF8) wrote a UTF-8 BOM followed by these two lines.
        private static readonly byte[] FileStart = BuildFileStart();

        private static byte[] BuildFileStart()
        {
            var nl = Environment.NewLine;
            var text = "From;TO" + nl + "FromConnector;FromPin;ToConnector;ToPin;Megger" + nl;
            var bom = Encoding.UTF8.GetPreamble();
            var body = Encoding.UTF8.GetBytes(text);
            var all = new byte[bom.Length + body.Length];
            Buffer.BlockCopy(bom, 0, all, 0, bom.Length);
            Buffer.BlockCopy(body, 0, all, bom.Length, body.Length);
            return all;
        }

        /// <summary>
        /// Writes {filePrefix}1.txt, {filePrefix}2.txt ... . The logical stream is
        /// lowLines (in list order) followed by the high pairs; every file holds maxLinesPerSheet
        /// lines except the last, and always at least one (header only) file exists - the same
        /// layout the original writer produced. Files are independent line ranges so they are
        /// written in parallel.
        /// </summary>
        public static List<string> WriteSheets(
            string folder,
            string filePrefix,
            IReadOnlyList<string> lowLines,
            MeggerHighPairs highPairs,
            long maxLinesPerSheet,
            int degreeOfParallelism)
        {
            long lowCount = lowLines?.Count ?? 0;
            long highCount = highPairs?.TotalLines ?? 0;
            long total = lowCount + highCount;
            int files = (int)Math.Max(1, (total + maxLinesPerSheet - 1) / maxLinesPerSheet);

            var paths = new string[files];
            for (int k = 0; k < files; k++)
            {
                paths[k] = Path.Combine(folder, $"{filePrefix}{k + 1}.txt");
            }

            Parallel.For(0, files,
                new ParallelOptions { MaxDegreeOfParallelism = Math.Max(1, degreeOfParallelism) },
                k =>
                {
                    long start = (long)k * maxLinesPerSheet;
                    long end = Math.Min(start + maxLinesPerSheet, total);

                    using var fs = new FileStream(paths[k], FileMode.Create, FileAccess.Write, FileShare.Read,
                        bufferSize: 1, FileOptions.SequentialScan);
                    using var sink = new MeggerByteSink(fs, SinkBufferSize);
                    sink.WriteBytes(FileStart);

                    for (long i = start; i < Math.Min(end, lowCount); i++)
                    {
                        sink.WriteTextLine(lowLines[(int)i]);
                    }

                    if (highPairs != null && end > lowCount)
                    {
                        highPairs.WriteRange(Math.Max(start, lowCount) - lowCount, end - lowCount, sink);
                    }

                    sink.Flush();
                });

            return new List<string>(paths);
        }

        /// <summary>Parallel file writers: disk bound, so a small default; override via MeggerWriteParallelism.</summary>
        public static int DefaultDegreeOfParallelism()
        {
            return Math.Clamp(Environment.ProcessorCount / 2, 1, 4);
        }
    }
}
