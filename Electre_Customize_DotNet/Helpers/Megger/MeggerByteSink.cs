#nullable disable
using System;
using System.IO;
using System.Text;

namespace Electre_Customize_DotNet.Helpers.Megger
{
    /// <summary>
    /// Buffered byte writer used for the Megger text files: no per-line string and no per-line
    /// stream call. Produces exactly what StreamWriter.WriteLine(UTF-8) would (platform newline).
    /// </summary>
    internal sealed class MeggerByteSink : IDisposable
    {
        private static readonly byte[] NewLine = Encoding.UTF8.GetBytes(Environment.NewLine);

        private readonly Stream _stream;
        private readonly byte[] _buf;
        private int _pos;

        public MeggerByteSink(Stream stream, int bufferSize)
        {
            _stream = stream;
            _buf = new byte[bufferSize];
        }

        public void WriteBytes(byte[] data)
        {
            if (data.Length > _buf.Length - _pos) Flush();
            if (data.Length > _buf.Length) { _stream.Write(data, 0, data.Length); return; }
            Buffer.BlockCopy(data, 0, _buf, _pos, data.Length);
            _pos += data.Length;
        }

        /// <summary>a ';' b suffix - the three parts of a high line.</summary>
        public void WriteLine(byte[] a, byte[] b, byte[] suffix)
        {
            int need = a.Length + 1 + b.Length + suffix.Length;
            if (need > _buf.Length - _pos) Flush();
            if (need > _buf.Length)
            {
                WriteBytes(a); WriteBytes(new[] { (byte)';' }); WriteBytes(b); WriteBytes(suffix);
                return;
            }
            Buffer.BlockCopy(a, 0, _buf, _pos, a.Length); _pos += a.Length;
            _buf[_pos++] = (byte)';';
            Buffer.BlockCopy(b, 0, _buf, _pos, b.Length); _pos += b.Length;
            Buffer.BlockCopy(suffix, 0, _buf, _pos, suffix.Length); _pos += suffix.Length;
        }

        /// <summary>A text line followed by the platform newline (what StreamWriter.WriteLine produced).</summary>
        public void WriteTextLine(string line)
        {
            int max = Encoding.UTF8.GetMaxByteCount(line.Length) + NewLine.Length;
            if (max > _buf.Length - _pos) Flush();
            if (max > _buf.Length)
            {
                WriteBytes(Encoding.UTF8.GetBytes(line));
                WriteBytes(NewLine);
                return;
            }
            _pos += Encoding.UTF8.GetBytes(line, 0, line.Length, _buf, _pos);
            Buffer.BlockCopy(NewLine, 0, _buf, _pos, NewLine.Length);
            _pos += NewLine.Length;
        }

        public void Flush()
        {
            if (_pos > 0)
            {
                _stream.Write(_buf, 0, _pos);
                _pos = 0;
            }
        }

        public void Dispose() => Flush();
    }
}
