namespace Electre_Customize_DotNet.Helpers.CableList
{
    internal static class CableListDataBlock
    {
        public static object[,] Build(object[,] iarr, int qStart, int rows)
        {
            object[,] block = new object[rows, 10];
            for (int r = 0; r < rows; r++)
            {
                int q = qStart + r;
                for (int p = 0; p < 9; p++)
                    block[r, p] = iarr[q, p + 1];
                block[r, 9] = iarr[q, 19];
            }
            return block;
        }
    }
}
