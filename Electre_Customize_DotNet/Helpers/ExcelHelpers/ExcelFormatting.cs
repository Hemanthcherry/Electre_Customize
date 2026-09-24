using Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace Electre_Customize_DotNet.Helpers.ExcelHelpers
{
    internal static class ExcelFormatting
    {
        /// <summary>Centered, wrapped, un-indented, unmerged - the standard look for report data blocks.</summary>
        public static void FormatRange(Range range)
        {
            range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            range.VerticalAlignment = XlVAlign.xlVAlignCenter;
            range.WrapText = true;
            range.Orientation = 0;
            range.AddIndent = false;
            range.IndentLevel = 0;
            range.ShrinkToFit = false;
            range.ReadingOrder = (int)Constants.xlContext;
            range.MergeCells = false;
        }
    }
}
