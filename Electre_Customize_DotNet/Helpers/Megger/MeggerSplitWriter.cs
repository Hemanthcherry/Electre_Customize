using Electre_Customize_DotNet.Logs;
using System.Text;

namespace Electre_Customize_DotNet.Helpers.Megger
{
    internal sealed class MeggerSplitWriter : IDisposable
    {
        internal const int BufferSize = 8 * 1024 * 1024;
        private const long MaxLinesPerFile = 5_000_000;
        private readonly string _folder;
        private readonly string _prefix;
        private StreamWriter? _writer;
        private int _sheet = 1;
        private long _linesInSheet;

        public long TotalLines { get; private set; }

        public MeggerSplitWriter(string folder, string prefix)
        {
            _folder = folder;
            _prefix = prefix;
            OpenSheet();
        }

        private void OpenSheet()
        {
            string path = Path.Combine(_folder, _prefix + "_" + _sheet + ".txt");
            var fs = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.Read, BufferSize, FileOptions.SequentialScan);
            _writer = new StreamWriter(fs, Encoding.UTF8, BufferSize)
            {
                AutoFlush = false,
                NewLine = "\r\n"
            };
            _writer.WriteLine("From;TO");
            _writer.WriteLine("FromConnector;FromPin;ToConnector;ToPin;Megger");
            _writer.Flush();
            _linesInSheet = 0;
        }

        private void Rotate()
        {
            Logging.Info(_prefix + "_" + _sheet + ".txt created successfully");
            _writer.Dispose();
            _sheet++;
            OpenSheet();
        }

        public void WriteExistingLine(string line)
        {
            if (_linesInSheet >= MaxLinesPerFile)
                Rotate();
            _writer.WriteLine(line);
            _linesInSheet++;
            TotalLines++;
        }

        public void AppendRawFile(string path)
        {
            _writer.Flush();
            using var src = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, BufferSize, FileOptions.SequentialScan);
            byte[] buf = new byte[BufferSize];
            int read;
            while ((read = src.Read(buf, 0, buf.Length)) > 0)
            {
                int chunkStart = 0;
                for (int i = 0; i < read; i++)
                {
                    if (buf[i] != (byte)'\n')
                        continue;
                    _linesInSheet++;
                    TotalLines++;
                    if (_linesInSheet < MaxLinesPerFile)
                        continue;
                    int len = i + 1 - chunkStart;
                    if (len > 0)
                        _writer.BaseStream.Write(buf, chunkStart, len);
                    Rotate();
                    chunkStart = i + 1;
                }
                if (chunkStart < read)
                    _writer.BaseStream.Write(buf, chunkStart, read - chunkStart);
            }
            _writer.Flush();
        }

        public void Dispose()
        {
            if (_writer == null)
                return;
            Logging.Info(_prefix + "_" + _sheet + ".txt created successfully");
            _writer.Dispose();
            _writer = null;
        }
    }
}
