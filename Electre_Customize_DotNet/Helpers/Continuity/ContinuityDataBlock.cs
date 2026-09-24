using Electre_Customize_DotNet.Helpers.Global;
using Microsoft.Office.Interop.Excel;

namespace Electre_Customize_DotNet.Helpers.Continuity
{
    internal static class ContinuityDataBlock
    {
        public static void Write(Worksheet ws, int startRow, object[,] arrFTcwob, int qStart, int rows)
        {
            object[,] colC = new object[rows, 1];
            object[,] colEF = new object[rows, 2];
            object[,] colHI = new object[rows, 2];
            object[,] colK = new object[rows, 1];
            object[,] colN = new object[rows, 1];

            for (int r = 0; r < rows; r++)
            {
                int q = qStart + r;
                colC[r, 0] = arrFTcwob[q, 0];
                colEF[r, 0] = arrFTcwob[q, 1];
                colEF[r, 1] = arrFTcwob[q, 2];
                colHI[r, 0] = arrFTcwob[q, 3];
                colHI[r, 1] = arrFTcwob[q, 4];
                colK[r, 0] = arrFTcwob[q, 5];
                colN[r, 0] = arrFTcwob[q, 6];
            }

            // Same cells as the original per-cell writes: C, E, F, H, I, K, N. Spacers D/G/J/L/M untouched.
            ExcelSheetOps.WriteBlock(ws, startRow, 3, colC);
            ExcelSheetOps.WriteBlock(ws, startRow, 5, colEF);
            ExcelSheetOps.WriteBlock(ws, startRow, 8, colHI);
            ExcelSheetOps.WriteBlock(ws, startRow, 11, colK);
            ExcelSheetOps.WriteBlock(ws, startRow, 14, colN);
        }
    }
}
