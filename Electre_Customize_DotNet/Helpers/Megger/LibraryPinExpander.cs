using System;
using System.Collections.Generic;
using System.Linq;

namespace Electre_Customize_DotNet.Helpers.Megger
{
    /// <summary>
    /// Expands the pin entries of a library connector (single pins and ranges such as "1-20", "A-Z", "A1-A5",
    /// "2A-5A") into the list of individual pin names. Moved unchanged out of GeneratePinListArray, which
    /// used to repeat this - including its repeated TakeWhile(char.IsLetter) scans - for every connector,
    /// although the result only depends on the (part number, gauge) pair.
    /// </summary>
    internal static class LibraryPinExpander
    {
        public static List<string> Expand(List<string> pinlistfromliabarary)
        {
            var expanded = new List<string>();

            foreach (string pin in pinlistfromliabarary)
            {
                if (pin.Contains("-"))
                {
                    string[] parts = pin.Split('-');
                    string start = parts[0];
                    string end = parts[1];

                    // Numeric range (1-20)
                    if (int.TryParse(start, out int startNum) && int.TryParse(end, out int endNum))
                    {
                        for (int num = startNum; num <= endNum; num++)
                        {
                            expanded.Add(num.ToString());
                        }
                    }
                    // Alphabetic range ( A-Z )
                    else if (start.Length == 1 && end.Length == 1 && char.IsLetter(start[0]) && char.IsLetter(end[0]))
                    {                                   
                        for (char c = start[0]; c <= end[0]; c++) 
                        { 
                            expanded.Add(c.ToString());
                        }
                    }
                    // Alphanumeric range with multiple letters prefix and number ranges (e.g., "A1-A5", "AA10-AA15", "AAA9-AAA19")
                    else if (start.Length > 1 && end.Length > 1 && start.TakeWhile(char.IsLetter).Count() == end.TakeWhile(char.IsLetter).Count() &&
                             start.TakeWhile(char.IsLetter).All(char.IsLetter) &&
                             int.TryParse(start.Substring(start.TakeWhile(char.IsLetter).Count()), out int startNumRange) &&
                             int.TryParse(end.Substring(end.TakeWhile(char.IsLetter).Count()), out int endNumRange) &&
                             start.Substring(0, start.TakeWhile(char.IsLetter).Count()) == end.Substring(0, start.TakeWhile(char.IsLetter).Count()))
                    {
                        // Extract the letter prefix (e.g., "A", "AA", "AAA")
                        string letterPrefix = start.Substring(0, start.TakeWhile(char.IsLetter).Count());

                        // Extract numeric parts of the start and end pins (e.g., 1 from "A1", 5 from "A5")
                        int startNumber = int.Parse(start.Substring(start.TakeWhile(char.IsLetter).Count()));
                        int endNumber = int.Parse(end.Substring(end.TakeWhile(char.IsLetter).Count()));

                        // Generate the range for the numbers
                        for (int num = startNumber; num <= endNumber; num++)
                        {
                            expanded.Add(letterPrefix + num.ToString());  // Combine the prefix and number
                        }
                    }
                    // Numeric range with multi-letter suffix (e.g., "2A-5A", "23AA-25AA", "227AAA-2343AAA")
                    else if (start.Length > 1 && end.Length > 1 &&
                             int.TryParse(start.Substring(0, start.Length - start.TakeWhile(char.IsLetter).Count()), out int startNum4) &&
                             int.TryParse(end.Substring(0, end.Length - end.TakeWhile(char.IsLetter).Count()), out int endNum4) &&
                             start.Substring(start.Length - start.TakeWhile(char.IsLetter).Count()) == end.Substring(end.Length - end.TakeWhile(char.IsLetter).Count()))
                    {
                        // Extract the numeric part and letter suffix part
                        string letterSuffix = start.Substring(start.Length - start.TakeWhile(char.IsLetter).Count());  // Extract the letter suffix (e.g., "AA")
                        int startNumber = int.Parse(start.Substring(0, start.Length - letterSuffix.Length));  // Extract the number part before the suffix
                        int endNumber = int.Parse(end.Substring(0, end.Length - letterSuffix.Length));  // Extract the number part before the suffix

                        // Loop through the number range and add the corresponding pin names
                        for (int num = startNumber; num <= endNumber; num++)
                        {
                            expanded.Add(num.ToString() + letterSuffix);  // Combine the number and the letter suffix
                        }
                    }

                    else if (start.Equals(end, StringComparison.OrdinalIgnoreCase))
                    {
                        expanded.Add(start);  // Just add the single pin
                    }
                    else
                    {
                        expanded.Add(pin);
                    }
                }
                else
                {
                    expanded.Add(pin);
                }
            }

            return expanded;
        }
    }
}
