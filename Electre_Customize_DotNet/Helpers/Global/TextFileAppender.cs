using System.Collections.Generic;
using System.IO;

namespace Electre_Customize_DotNet.Helpers.Global
{
    internal static class TextFileAppender
    {
        /// <summary>Appends each line (in order) to the file, creating it if needed.</summary>
        public static void AppendLines(string filePath, List<string> lines)
        {
            using (StreamWriter writer = new StreamWriter(filePath, append: true))
            {
                foreach (string line in lines)
                {
                    writer.WriteLine(line);
                }
            }
        }
    }
}
