namespace Electre_Customize_DotNet.Objects
{
    /// <summary>
    /// EQU left/right connector suffixes (J1↔a … J24↔z, skipping i and o).
    /// </summary>
    internal static class EquConnectorMap
    {
        public static readonly Dictionary<string, string> Pairs =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "J1", "a" }, { "a", "J1" },
                { "J2", "b" }, { "b", "J2" },
                { "J3", "c" }, { "c", "J3" },
                { "J4", "d" }, { "d", "J4" },
                { "J5", "e" }, { "e", "J5" },
                { "J6", "f" }, { "f", "J6" },
                { "J7", "g" }, { "g", "J7" },
                { "J8", "h" }, { "h", "J8" },
                { "J9", "j" }, { "j", "J9" },
                { "J10", "k" }, { "k", "J10" },
                { "J11", "l" }, { "l", "J11" },
                { "J12", "m" }, { "m", "J12" },
                { "J13", "n" }, { "n", "J13" },
                { "J14", "p" }, { "p", "J14" },
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

        public static readonly HashSet<string> MaleSuffixes = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "J1", "J2", "J3", "J4", "J5", "J6", "J7", "J8",
            "J9", "J10", "J11", "J12", "J13", "J14", "J15", "J16",
            "J17", "J18", "J19", "J20", "J21", "J22", "J23", "J24"
        };

        public static readonly HashSet<string> FemaleSuffixes = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "a", "b", "c", "d", "e", "f", "g", "h",
            "j", "k", "l", "m", "n", "p", "q", "r",
            "s", "t", "u", "v", "w", "x", "y", "z"
        };
    }
}
