using Electre_Customize_DotNet.Helpers.Global;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace Electre_Customize_DotNet.Helpers.CableList
{
    internal static class CableListStrikethrough
    {
        public static void Apply(Application app, Worksheet ws, List<int> excelRows)
        {
            if (excelRows == null || excelRows.Count == 0)
                return;

            Range? union = null;
            int batch = 0;
            for (int i = 0; i < excelRows.Count; i++)
            {
                Range rowRange = ws.Range["B" + excelRows[i], "L" + excelRows[i]];
                union = union == null ? rowRange : app.Union(union, rowRange);
                batch++;
                if (batch == 20)
                {
                    union.MergeCells = false;
                    union.Font.Strikethrough = true;
                    ExcelRangeHelper.ReleaseCom(union);
                    union = null;
                    batch = 0;
                }
            }
            if (union != null)
            {
                union.MergeCells = false;
                union.Font.Strikethrough = true;
                ExcelRangeHelper.ReleaseCom(union);
            }
        }
    }
}
