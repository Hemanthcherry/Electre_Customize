using Electre_Customize_DotNet.Logs;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace Electre_Customize_DotNet.Helpers.Global
{
    internal static class ExcelSheetOps
    {
        public static void WriteBlock(Worksheet ws, int startRow, int startCol, object[,] block)
        {
            if (ws == null || block == null)
                return;

            int rows = block.GetLength(0);
            int cols = block.GetLength(1);
            if (rows == 0 || cols == 0)
                return;

            Range dest = ws.Range[
                ExcelRangeHelper.ToA1(startRow, startCol) + ":" +
                ExcelRangeHelper.ToA1(startRow + rows - 1, startCol + cols - 1)];
            dest.Value2 = block;
            ExcelRangeHelper.ReleaseCom(dest);
        }

        public static void Ungroup(Workbook workbook)
        {
            if (workbook == null || workbook.Sheets.Count < 1)
                return;
            try { ((Worksheet)workbook.Sheets[1]).Select(Type.Missing); } catch { }
        }

        public static void SelectGroup(Workbook workbook, int fromSheet, int toSheet)
        {
            ((Worksheet)workbook.Sheets[fromSheet]).Select(Type.Missing);
            for (int i = fromSheet + 1; i <= toSheet; i++)
                ((Worksheet)workbook.Sheets[i]).Select(false);
        }

        public static void ClearCutCopyMode(Application app)
        {
            try { app.CutCopyMode = 0; } catch { }
        }

        public static Worksheet? FindByName(Workbook workbook, string sheetName, bool ignoreCase = true)
        {
            if (workbook == null || string.IsNullOrEmpty(sheetName))
                return null;

            try
            {
                Worksheet exact = (Worksheet)workbook.Sheets[sheetName];
                if (exact != null)
                    return exact;
            }
            catch { }

            if (!ignoreCase)
                return null;

            foreach (Worksheet sheet in workbook.Sheets)
            {
                if (string.Equals(sheet.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                    return sheet;
            }
            return null;
        }

        public static void KeepOnlyNamedSheet(Workbook workbook, Worksheet keep)
        {
            string keepName = keep.Name;
            for (int i = workbook.Sheets.Count; i >= 1; i--)
            {
                Worksheet ws = (Worksheet)workbook.Sheets[i];
                if (!string.Equals(ws.Name, keepName, StringComparison.OrdinalIgnoreCase))
                    ws.Delete();
            }
        }

        public static void DuplicateTemplateSheets(Application app, Workbook workbook, int needed, string namePrefix)
        {
            if (needed <= 1)
            {
                Worksheet only = (Worksheet)workbook.Sheets[1];
                if (only.Name != namePrefix + "1")
                    only.Name = namePrefix + "1";
                return;
            }

            app.DisplayAlerts = false;
            Ungroup(workbook);

            Worksheet template = (Worksheet)workbook.Sheets[1];
            if (template.Name != namePrefix + "1")
                template.Name = namePrefix + "1";

            // One-sheet copies of the original HAL page. In Open XML (.xlsx) this is faster
            // than doubling (which copies 32–64 heavy sheets in one COM call).
            for (int i = 2; i <= needed; i++)
            {
                template.Copy(After: workbook.Sheets[workbook.Sheets.Count]);
                ((Worksheet)workbook.Sheets[workbook.Sheets.Count]).Name = namePrefix + i;
            }

            Ungroup(workbook);
        }

        public static void CopyRangeToSheetGroup(Application app, Workbook workbook, Worksheet source, string address, int fromSheet, int toSheet)
        {
            if (source == null || fromSheet > toSheet)
                return;

            SelectGroup(workbook, fromSheet, toSheet);
            Worksheet dest = (Worksheet)app.ActiveSheet;
            source.Range[address].Copy(dest.Range[address]);
            Ungroup(workbook);
            ClearCutCopyMode(app);
        }
    }
}
