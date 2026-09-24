using System;
using System.Configuration;
using Microsoft.Office.Interop.Excel;

namespace Electre_Customize_DotNet.Helpers.ExcelHelpers
{
    /// <summary>
    /// Creates the numbered report sheets (CL-1, CL-2 ... / CWOB-1, CWOB-2 ...) from the template sheet.
    ///
    /// Default (App.config IntraWorkbookSheetCopy = false): every sheet is copied from the template workbook,
    /// exactly as before. Copying between two workbooks re-imports the template's styles each time, which is
    /// usually the slowest single step of a Cable List / Continuity job with many sheets.
    ///
    /// Opt-in (IntraWorkbookSheetCopy = true): only the FIRST sheet is copied from the template; the others are
    /// copied from that first (still untouched) sheet inside the report workbook, which is much cheaper. At that
    /// point no data has been written yet, so the copies should be identical - but that has to be confirmed on
    /// real drawings, which is why it is off by default.
    /// </summary>
    internal static class TemplateSheetCopier
    {
        public static bool IntraWorkbookCopyEnabled()
        {
            return string.Equals(ConfigurationManager.AppSettings["IntraWorkbookSheetCopy"], "true", StringComparison.OrdinalIgnoreCase);
        }

        public static void CopyInto(Workbook target, Worksheet template, int count, string namePrefix)
        {
            bool intraWorkbook = IntraWorkbookCopyEnabled();
            Worksheet firstCopy = null;

            for (int l = 1; l <= count; l++)
            {
                // Copy the template sheet (or, in opt-in mode, the first copy) to the report workbook
                Worksheet source = (intraWorkbook && firstCopy != null) ? firstCopy : template;
                source.Copy(After: target.Sheets[target.Sheets.Count]);

                // Rename the newly copied sheet
                Worksheet newSheet = target.Sheets[target.Sheets.Count];
                newSheet.Name = namePrefix + l;

                if (firstCopy == null)
                {
                    firstCopy = newSheet;
                }
            }
        }
    }
}
