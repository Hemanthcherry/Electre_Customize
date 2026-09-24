using System;
using System.Collections.Generic;
using System.Linq;
using Electre_Customize_DotNet.Logs;
using Electre_Customize_DotNet.Objects;

namespace Electre_Customize_DotNet.Helpers.Megger
{
    /// <summary>Conversions between the array/dictionary shapes the Megger and Continuity sheets exchange.</summary>
    internal static class MeggerListConverters
    {
        /// <summary>Continuity rows (connector1, pin1, connector2, pin2) -> "c1;p1;c2;p2;Low Megger" lines. Rows with any blank column are skipped.</summary>
        public static List<string> ToLowMeggerLines(string[,] arrContList)
        {
            var result = new List<string>();

            int rows = arrContList.GetLength(0);

            for (int i = 0; i < rows; i++)
            {
                string connector1 = arrContList[i, 0];
                string pin1 = arrContList[i, 1];
                string connector2 = arrContList[i, 2];
                string pin2 = arrContList[i, 3];
                if (string.IsNullOrWhiteSpace(connector1) ||
                    string.IsNullOrWhiteSpace(pin1) ||
                    string.IsNullOrWhiteSpace(connector2) ||
                    string.IsNullOrWhiteSpace(pin2))
                {
                    continue;
                }
                string resultLine = $"{connector1};{pin1};{connector2};{pin2};Low Megger";
                result.Add(resultLine);
            }

            return result;
        }

        /// <summary>Converts data from Dictionary to a 2D array for the Megger custom-loom sheet (key in column 0, values after it).</summary>
        public static object[,] ToWideArray(Dictionary<string, List<string>> dict)
        {
            try
            {

                int rows = dict.Count;
                int cols = dict.Values.Max(list => list.Count) + 1; // +1 for key column

                object[,] array2D = new object[rows, cols];

                int row = 0;
                foreach (var kvp in dict)
                {
                    array2D[row, 0] = kvp.Key; // First column for keys
                    for (int col = 0; col < kvp.Value.Count; col++)
                    {
                        array2D[row, col + 1] = kvp.Value[col]; // Remaining columns for values
                    }
                    row++;
                }
                return array2D;
            }
            catch (Exception ex)
            {

                // Log error message
                Logging.Error("modMain: #009 " + ex.Message);


            }
            return new object[0, 0];  // Return empty array on error
        }

        /// <summary>Converts data from a 2D array to a Dictionary (column 0 is the key, non-null cells after it are the values).</summary>
        public static Dictionary<string, List<string>> ToDictionary(object[,] array2D)
        {
            Dictionary<string, List<string>> dict = new Dictionary<string, List<string>>();

            for (int i = 0; i < array2D.GetLength(0); i++)
            {
                string key = array2D[i, 0]?.ToString(); // First column as key
                List<string> values = new List<string>();

                for (int j = 1; j < array2D.GetLength(1); j++) // Remaining columns as values
                {
                    if (array2D[i, j] != null) // Ignore null values
                    {
                        values.Add(array2D[i, j].ToString());
                    }
                }

                dict[key] = values;
            }
            return dict;
        }

        /// <summary>Serial number, connector name and pin for each object that has no Megger data.</summary>
        public static string[,] NoMeggerToArray(List<ElectreObject> noMeggerData)
        {
            if (noMeggerData == null || noMeggerData.Count == 0)
            {
                return new string[0, 0]; // Return empty array if no data
            }

            int rowCount = noMeggerData.Count;
            string[,] result = new string[rowCount, 3]; // 3 columns: Serial, ConnectorName, PinNumber

            for (int i = 0; i < rowCount; i++)
            {
                var obj = noMeggerData[i];

                result[i, 0] = (i + 1).ToString();             // Serial Number (starts from 1)
                result[i, 1] = obj.ConnectorName ?? "";        // Safe null-check
                result[i, 2] = obj.PinNumber ?? "";
            }

            return result;
        }
    }
}
