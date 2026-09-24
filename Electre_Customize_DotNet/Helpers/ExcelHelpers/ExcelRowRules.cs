using System;

namespace Electre_Customize_DotNet.Helpers.ExcelHelpers
{
    /// <summary>Row/paging rules applied to the in-memory report arrays before they are written to a sheet.</summary>
    internal static class ExcelRowRules
    {
        /// <summary>True when the row and the next two rows are blank in columns 1 to 4.</summary>
        public static bool IsBlockBlank(object[,] iarr, int currentRow)
        {
            try
            {
                // Check if the current row and the next two rows are blank in columns 1 to 4
                return IsRowBlank(iarr, currentRow) &&
                       IsRowBlank(iarr, currentRow + 1) &&
                       IsRowBlank(iarr, currentRow + 2);
            }
            catch
            {
                return false;
            }
        }

        public static bool IsRowBlank(object[,] iarr, int row)
        {
            // Check if all cells in columns 1 to 4 of the specified row are blank
            return row < iarr.GetLength(0) &&
                   string.IsNullOrEmpty(iarr[row, 0]?.ToString()) &&
                   string.IsNullOrEmpty(iarr[row, 1]?.ToString()) &&
                   string.IsNullOrEmpty(iarr[row, 2]?.ToString()) &&
                   string.IsNullOrEmpty(iarr[row, 3]?.ToString());
        }

        public static int GetNumberOfSheetsRequired(int iTotalLines, int iNoLinesCanBePrinted)
        {
            double X = (double)iTotalLines / iNoLinesCanBePrinted;

            if (X == Math.Round(X, 0))
            {
                return (int)X;
            }
            else
            {
                return (int)Math.Round(X + 0.5, 0);
            }
        }

        /// <summary>Last populated row index before the first block of three blank rows (columns 1-4).</summary>
        public static int MaxRowinArray(object[,] iarr)
        {
            if (iarr == null)
                throw new ArgumentNullException(nameof(iarr));

            int lowerBoundRow = iarr.GetLowerBound(0); // usually 0 or 1
            int upperBoundRow = iarr.GetUpperBound(0); // max row index
            int lowerBoundCol = iarr.GetLowerBound(1); // usually 0 or 1
            int upperBoundCol = iarr.GetUpperBound(1);

            for (int K = lowerBoundRow; K <= upperBoundRow; K++)
            {
                bool isCurrentAndNext2RowsEmpty = true;

                // Check only if we have 3 rows ahead
                if (K + 2 <= upperBoundRow)
                {
                    for (int i = 0; i < 3; i++) // Check K, K+1, K+2
                    {
                        for (int j = 1; j <= 4; j++) // Columns 1 to 4
                        {
                            if (iarr[K + i, j] != null && iarr[K + i, j].ToString().Trim() != "")
                            {
                                isCurrentAndNext2RowsEmpty = false;
                                break;
                            }
                        }
                        if (!isCurrentAndNext2RowsEmpty)
                            break;
                    }

                    if (isCurrentAndNext2RowsEmpty)
                        return K - 1;
                }
            }

            // If we reached the end of the loop without breaking early, return the last row
            return upperBoundRow;
        }
    }
}
