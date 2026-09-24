using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace Electre_Customize_DotNet.Helpers.ExcelHelpers
{
    /// <summary>
    /// Whole-block Excel reads and writes. Every COM call crosses a process boundary, so the report
    /// generators build an object[,] in memory and move it in ONE call instead of touching cells one
    /// at a time. The temporary Range COM object is always released.
    /// </summary>
    internal static class ExcelBulk
    {
        /// <summary>
        /// Writes <paramref name="values"/> (0-based) with its top-left cell at (firstRow, firstCol).
        /// Uses Range.Value (not Value2) - the original per-cell writes used Value, and the two differ in
        /// how dates and currency-like text are interpreted.
        /// </summary>
        public static void WriteValues(Excel.Worksheet sheet, int firstRow, int firstCol, object[,] values, bool centered = false)
        {
            int lastRow = firstRow + values.GetLength(0) - 1;
            int lastCol = firstCol + values.GetLength(1) - 1;

            // One "C9:L9" style address = one COM call to get the range; sheet.Range[sheet.Cells[..], sheet.Cells[..]]
            // costs five (Cells, indexer, Cells, indexer, Range) - and this runs once per report row.
            Excel.Range range = RangeAt(sheet, firstRow, firstCol, lastRow, lastCol);
            try
            {
                range.Value = values;
                if (centered)
                {
                    range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                }
            }
            finally
            {
                Marshal.ReleaseComObject(range);
            }
        }

        /// <summary>
        /// The rectangle (firstRow, firstCol)..(lastRow, lastCol) as ONE COM call. Equivalent to
        /// sheet.Range[sheet.Cells[firstRow, firstCol], sheet.Cells[lastRow, lastCol]], which costs five.
        /// </summary>
        public static Excel.Range RangeAt(Excel.Worksheet sheet, int firstRow, int firstCol, int lastRow, int lastCol)
        {
            return sheet.Range[ColumnLetters(firstCol) + firstRow + ":" + ColumnLetters(lastCol) + lastRow];
        }

        /// <summary>1 -> "A", 26 -> "Z", 27 -> "AA" ...</summary>
        private static string ColumnLetters(int column)
        {
            string letters = string.Empty;
            while (column > 0)
            {
                int remainder = (column - 1) % 26;
                letters = (char)('A' + remainder) + letters;
                column = (column - 1) / 26;
            }
            return letters;
        }

        /// <summary>
        /// Reads the rectangle in one call. For a multi-cell range Excel returns a 1-based object[,], so
        /// result[row, col] lines up with (firstRow + row - 1, firstCol + col - 1). Callers must not ask
        /// for a single cell (Excel returns a scalar there).
        /// </summary>
        public static object[,] ReadValues(Excel.Worksheet sheet, int firstRow, int firstCol, int lastRow, int lastCol)
        {
            Excel.Range range = sheet.Range[sheet.Cells[firstRow, firstCol], sheet.Cells[lastRow, lastCol]];
            try
            {
                return (object[,])range.Value2;
            }
            finally
            {
                Marshal.ReleaseComObject(range);
            }
        }
    }
}
