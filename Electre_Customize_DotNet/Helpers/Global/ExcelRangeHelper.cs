using System.Runtime.InteropServices;

namespace Electre_Customize_DotNet.Helpers.Global
{
    internal static class ExcelRangeHelper
    {
        public static void ReleaseCom(object comObj)
        {
            if (comObj != null && Marshal.IsComObject(comObj))
            {
                try { Marshal.ReleaseComObject(comObj); } catch { }
            }
        }

        public static string ToA1(int row, int col)
        {
            int c = col;
            string colName = string.Empty;
            while (c > 0)
            {
                int rem = (c - 1) % 26;
                colName = (char)('A' + rem) + colName;
                c = (c - 1) / 26;
            }
            return colName + row;
        }
    }
}
