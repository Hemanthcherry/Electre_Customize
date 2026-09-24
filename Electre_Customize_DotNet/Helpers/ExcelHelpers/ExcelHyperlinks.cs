using System;
using System.IO;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace Electre_Customize_DotNet.Helpers.ExcelHelpers
{
    internal static class ExcelHyperlinks
    {
        /// <summary>Lists every "{prefix}*.txt" file in the folder as a yellow, bold hyperlink down column A.</summary>
        public static void AddTextFileLinks(Worksheet sheet, string folderPath, string filePrefix)
        {
            string[] filePaths = Directory.GetFiles(folderPath, $"{filePrefix}*.txt");
            int row = 1;
            bool widthSet = false;

            foreach (string file in filePaths.OrderBy(f => f))
            {
                string fileName = Path.GetFileName(file);
                string displayText = $"Refer to {fileName} - Click to Open";

                Range hyperlinkCell = ExcelBulk.RangeAt(sheet, row, 1, row, 1); // Column A
                if (!widthSet)
                {
                    sheet.Columns[1].ColumnWidth = 40;   // same value every time - set it once
                    widthSet = true;
                }

                sheet.Hyperlinks.Add(hyperlinkCell, file, Type.Missing, Type.Missing, displayText);
                hyperlinkCell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                hyperlinkCell.Font.Bold = true;

                row++;
            }
        }
    }
}
