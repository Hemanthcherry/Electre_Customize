namespace Electre_Customize_DotNet.Objects
{
    /// <summary>
    /// Cable-type families used to build wire codes.
    /// Matching is the original joined-string Contains("_" + type + "_"),
    /// including empty/null types which match the "_" sentinels.
    /// </summary>
    internal static class CableTypes
    {
        private static readonly string[] NormalItems =
        {
            "_", "PT", "PB", "PTB", "TB", "TTB", "QT", "QB", "QTB", "HTTP", "HTTQ", "HTTT",
            "HTSTP", "HTSTQ", "HTSTT", "SP", "TP", "STP", "ST", "TT", "STT", "SQ", "TQ",
            "STQ", "HTS", "HTSS", "ST5", "ST6", "ST7", "ST8", "ST9", "ST10", "_"
        };

        private static readonly string[] Special1Items = { "_", "X", "TX", "BX", "_" };

        private static readonly string[] Special2Items = { "_", "CAT5STP", "CAT5STQ", "BC", "_" };

        public static readonly string JoinedNormal = string.Join("_", NormalItems);
        public static readonly string JoinedSpecial1 = string.Join("_", Special1Items);
        public static readonly string JoinedSpecial2 = string.Join("_", Special2Items);

        public static bool IsNormalFamily(string cableType)
        {
            return JoinedNormal.Contains("_" + cableType + "_");
        }

        public static bool IsSpecial1(string cableType)
        {
            return JoinedSpecial1.Contains("_" + cableType + "_");
        }

        public static bool IsSpecial2(string cableType)
        {
            return JoinedSpecial2.Contains("_" + cableType + "_");
        }
    }
}
