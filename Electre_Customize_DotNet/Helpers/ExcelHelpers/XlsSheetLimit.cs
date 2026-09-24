using System;
using System.Configuration;
using Electre_Customize_DotNet.Logs;
using Excel = Microsoft.Office.Interop.Excel;

namespace Electre_Customize_DotNet.Helpers.ExcelHelpers
{
    /// <summary>
    /// The .xls (BIFF8) file format cannot hold more than 255 sheets in one workbook - Excel refuses to add the
    /// 256th sheet, so a Cable List / Continuity report that needs more sheets than that fails outright partway
    /// through the sheet-copy loop, with no usable file produced at all. Below the limit nothing here changes
    /// anything: same ".xls", same XlFileFormat.xlExcel8, byte-for-byte the original behaviour.
    ///
    /// Merged in from a review of another optimisation pass (Grok's) that used this exact fix for this exact
    /// problem - the one idea from that pass that didn't touch calculation mode, sort order or row-truncation
    /// logic, so it carries none of that pass's correctness risk.
    /// </summary>
    internal static class XlsSheetLimit
    {
        /// <summary>The real Excel limit: a .xls workbook can hold at most this many sheets.</summary>
        public const int MaxSheetsPerXls = 255;

        /// <summary>
        /// Default ON. Below the limit this makes no difference. Above it, the only alternative is Excel throwing a
        /// COM exception mid-copy and no report being produced - there is no existing correct .xls behaviour above
        /// 255 sheets to preserve, so defaulting to "produce a working .xlsx instead of no file at all" is the safer
        /// default. Set to "false" to keep the old all-or-nothing .xls behaviour (report simply fails above the limit).
        /// </summary>
        public static bool AutoXlsxAboveLimitEnabled()
        {
            var raw = ConfigurationManager.AppSettings["AutoXlsxAboveSheetLimit"];
            return string.IsNullOrEmpty(raw) || string.Equals(raw, "true", StringComparison.OrdinalIgnoreCase);
        }

        public readonly struct ResolvedFormat
        {
            public readonly string FileName;
            public readonly Excel.XlFileFormat FileFormat;
            public readonly bool UsingXlsx;

            public ResolvedFormat(string fileName, Excel.XlFileFormat fileFormat, bool usingXlsx)
            {
                FileName = fileName;
                FileFormat = fileFormat;
                UsingXlsx = usingXlsx;
            }
        }

        /// <summary>Decides the file name and Excel save format for a report that needs <paramref name="numberOfSheetsRequired"/> sheets.</summary>
        public static ResolvedFormat ResolveFormat(string reportName, int numberOfSheetsRequired)
        {
            bool useXlsx = numberOfSheetsRequired > MaxSheetsPerXls && AutoXlsxAboveLimitEnabled();
            if (useXlsx)
            {
                Logging.Info($"'{reportName}' needs {numberOfSheetsRequired} sheets; .xls holds at most {MaxSheetsPerXls} sheets, saving as .xlsx instead.");
            }
            return new ResolvedFormat(
                reportName + (useXlsx ? ".xlsx" : ".xls"),
                useXlsx ? Excel.XlFileFormat.xlOpenXMLWorkbook : Excel.XlFileFormat.xlExcel8,
                useXlsx);
        }
    }
}
