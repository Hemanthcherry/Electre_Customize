using System;

namespace Electre_Customize_DotNet.Helpers.CableList
{
    /// <summary>
    /// The per-sheet / per-component report buffers used to be allocated at the full height of the
    /// project array (every row, for every sheet), although only the rows of one sheet were ever filled.
    /// The writer (AppendToExcel) stops at the first block of three blank rows, so the buffer only needs
    /// its filled rows plus a few blank rows behind them - BUT it must still be taller than three rows,
    /// because AppendToExcel only looks for the blank block when the array has more than 3 rows.
    /// </summary>
    internal static class RowBuffer
    {
        /// <summary>Filled rows are followed by this many blank ones (3 form the stop block, +1 keeps the height above 3).</summary>
        public const int TrailingBlankRows = 4;

        /// <summary>
        /// Height of the smallest buffer that makes AppendToExcel write exactly what a buffer of <paramref name="fullHeight"/> rows
        /// (the first rows of <paramref name="source"/>, the rest empty) would: everything up to and including the first row that starts
        /// a block of three blank rows (columns 1-4), plus the blank rows that prove it. Used where the original allocated
        /// rows x columns rows for a single big list.
        /// </summary>
        public static int EquivalentHeight(object[,] source, int fullHeight)
        {
            if (fullHeight <= 3)
            {
                return fullHeight;   // AppendToExcel only looks for the blank block when the height is above 3
            }

            int sourceRows = source.GetLength(0);

            bool IsBlank(int row)
            {
                if (row >= fullHeight)
                {
                    return false;    // IsRowBlank is false outside the array
                }
                if (row >= sourceRows)
                {
                    return true;     // rows the original left empty
                }
                return string.IsNullOrEmpty(source[row, 0]?.ToString()) &&
                       string.IsNullOrEmpty(source[row, 1]?.ToString()) &&
                       string.IsNullOrEmpty(source[row, 2]?.ToString()) &&
                       string.IsNullOrEmpty(source[row, 3]?.ToString());
            }

            for (int k = 0; k < fullHeight; k++)
            {
                if (IsBlank(k) && IsBlank(k + 1) && IsBlank(k + 2))
                {
                    return Math.Min(fullHeight, k + TrailingBlankRows);
                }
            }
            return fullHeight;
        }

        /// <summary>
        /// Grows <paramref name="buffer"/> (keeping its content) to <c>min(maxRows, filledRows + TrailingBlankRows)</c>
        /// rows if it is shorter. <paramref name="maxRows"/> is the height the original always allocated,
        /// so the buffer never ends up taller than before.
        /// </summary>
        public static void EnsureCapacity(ref object[,] buffer, int filledRows, int maxRows)
        {
            int columns = buffer.GetLength(1);
            int wanted = Math.Min(maxRows, filledRows + TrailingBlankRows);
            if (buffer.GetLength(0) >= wanted)
            {
                return;
            }

            var grown = new object[wanted, columns];
            Array.Copy(buffer, grown, buffer.Length); // same column count, so a row-major copy keeps every cell in place
            buffer = grown;
        }
    }

    /// <summary>The reusable row buffer and fill counter of one worker of the sheet-wise wire-list report.</summary>
    internal sealed class WirelistBufferState
    {
        public object[,] Rows;
        public int Z;
        /// <summary>Height cap for the buffer and the sort request (the component report uses a different one for a worker's first item).</summary>
        public int MaxRows;

        public WirelistBufferState(int columns, int maxRows = int.MaxValue)
        {
            Rows = new object[0, columns];
            MaxRows = maxRows;
        }
    }
}
