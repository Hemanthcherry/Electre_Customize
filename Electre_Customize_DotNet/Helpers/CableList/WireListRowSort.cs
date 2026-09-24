namespace Electre_Customize_DotNet.Helpers.CableList
{
    internal static class WireListRowSort
    {
        public static void SortByWireCode(object[,] rows, int rowCount, int wireCodeCol = 4)
        {
            if (rows == null || rowCount < 2)
                return;

            int cols = rows.GetLength(1);
            var order = new int[rowCount];
            for (int i = 0; i < rowCount; i++)
                order[i] = i;

            Array.Sort(order, (a, b) =>
            {
                string wa = rows[a, wireCodeCol]?.ToString() ?? "";
                string wb = rows[b, wireCodeCol]?.ToString() ?? "";
                return string.Compare(wa, wb, StringComparison.OrdinalIgnoreCase);
            });

            bool identity = true;
            for (int i = 0; i < rowCount; i++)
            {
                if (order[i] != i)
                {
                    identity = false;
                    break;
                }
            }
            if (identity)
                return;

            object[,] copy = new object[rowCount, cols];
            for (int r = 0; r < rowCount; r++)
            {
                int src = order[r];
                for (int c = 0; c < cols; c++)
                    copy[r, c] = rows[src, c];
            }
            for (int r = 0; r < rowCount; r++)
                for (int c = 0; c < cols; c++)
                    rows[r, c] = copy[r, c];
        }
    }
}
