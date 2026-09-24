using System.IO;

namespace Electre_Customize_DotNet.Helpers.Global
{
    internal static class SheetNames
    {
        /// <summary>Strips characters that are invalid in file/sheet names and trims to Excel's 31-character limit.</summary>
        public static string Sanitize(string sheetName)
        {
            // Remove invalid characters
            string invalidChars = new string(Path.GetInvalidFileNameChars()) + @":\/?*[]";
            foreach (char c in invalidChars)
            {
                sheetName = sheetName.Replace(c.ToString(), "");
            }

            // Trim to 31 characters
            return sheetName.Length > 31 ? sheetName.Substring(0, 31) : sheetName;
        }
    }
}
