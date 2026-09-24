using Electre_Customize_DotNet.Objects;

namespace Electre_Customize_DotNet.Helpers.Global
{
    internal static class PairWireCode
    {
        public static string Build(ElectreObject E1, bool extraSlashBeforeCore)
        {
            string coreNumber = IsValidCoreNumber(E1.Core_Part_Number) ? E1.Core_Part_Number : "";
            string gaugeMid = (!string.IsNullOrEmpty(E1.Gauge) && E1.Gauge.Length > 1) ? E1.Gauge.Substring(1) : "";

            if (!(CableTypes.IsNormalFamily(E1.CableType) && E1.CableType != "HTSS")
                && (CableTypes.IsSpecial1(E1.CableType) || CableTypes.IsSpecial2(E1.CableType)))
            {
                return $"{E1.WireNumber}/{coreNumber}";
            }

            if (coreNumber == "" && E1.CableType != "X")
                return $"{E1.WireNumber}/{gaugeMid}";

            if (extraSlashBeforeCore)
                return $"{E1.WireNumber}/{(gaugeMid.Length > 0 ? gaugeMid + "/" : "")}/{coreNumber}";

            return $"{E1.WireNumber}/{(gaugeMid.Length > 0 ? gaugeMid + "/" : "")}{coreNumber}";
        }

        public static bool IsValidCoreNumber(string tag3)
        {
            return int.TryParse(tag3, out int coreNumber) && coreNumber >= 1 && coreNumber <= 18;
        }
    }
}
