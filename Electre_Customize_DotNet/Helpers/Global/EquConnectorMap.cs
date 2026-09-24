using System;
using System.Collections.Generic;

namespace Electre_Customize_DotNet.Helpers.Global
{
    /// <summary>
    /// Pairs used for swapping the EQU connectors (J1 &lt;-&gt; a, J2 &lt;-&gt; b, ...). Lookup is
    /// case-insensitive. Callers only read it (Keys / indexer / enumeration), never modify it.
    /// </summary>
    internal static class EquConnectorMap
    {
        public static readonly Dictionary<string, string> Pairs = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "J1", "a" }, { "a", "J1" },
            { "J2", "b" }, { "b", "J2" },
            { "J3", "c" }, { "c", "J3" },
            { "J4", "d" }, { "d", "J4" },
            { "J5", "e" }, { "e", "J5" },
            { "J6", "f" }, { "f", "J6" },
            { "J7", "g" }, { "g", "J7" },
            { "J8", "h" }, { "h", "J8" },
            { "J9", "j" }, { "j", "J9" },   // Skipped 'i'
            { "J10", "k" }, { "k", "J10" },
            { "J11", "l" }, { "l", "J11" },
            { "J12", "m" }, { "m", "J12" },
            { "J13", "n" }, { "n", "J13" },
            { "J14", "p" }, { "p", "J14" }, // Skipped 'o'
            { "J15", "q" }, { "q", "J15" },
            { "J16", "r" }, { "r", "J16" },
            { "J17", "s" }, { "s", "J17" },
            { "J18", "t" }, { "t", "J18" },
            { "J19", "u" }, { "u", "J19" },
            { "J20", "v" }, { "v", "J20" },
            { "J21", "w" }, { "w", "J21" },
            { "J22", "x" }, { "x", "J22" },
            { "J23", "y" }, { "y", "J23" },
            { "J24", "z" }, { "z", "J24" },
        };

        // Snapshots in the dictionary's own (insertion) enumeration order, with the "_" prefix the trace and sort code
        // tests for. Precomputed once: the callers used to build "_" + suffix for all 48 keys for every connector.
        private static readonly string[] Suffixes = new List<string>(Pairs.Keys).ToArray();
        private static readonly string[] UnderscoredSuffixes = Array.ConvertAll(Suffixes, k => "_" + k);

        /// <summary>Same as <c>Pairs.Keys.Any(k => connectorName.EndsWith("_" + k, OrdinalIgnoreCase))</c>.</summary>
        public static bool EndsWithUnderscoredKey(string connectorName)
        {
            for (int i = 0; i < UnderscoredSuffixes.Length; i++)
            {
                if (connectorName.EndsWith(UnderscoredSuffixes[i], StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>Same as <c>Pairs.Keys.FirstOrDefault(k => connectorName.EndsWith("_" + k, OrdinalIgnoreCase))</c> (null if none).</summary>
        public static string FirstUnderscoredKey(string connectorName)
        {
            for (int i = 0; i < UnderscoredSuffixes.Length; i++)
            {
                if (connectorName.EndsWith(UnderscoredSuffixes[i], StringComparison.OrdinalIgnoreCase))
                {
                    return Suffixes[i];
                }
            }
            return null;
        }
    }
}
