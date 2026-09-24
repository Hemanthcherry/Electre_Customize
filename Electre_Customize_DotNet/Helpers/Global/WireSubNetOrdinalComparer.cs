namespace Electre_Customize_DotNet.Helpers.Global
{
    internal sealed class WireSubNetOrdinalComparer : IEqualityComparer<(string Wire, string SubNet)>
    {
        public static readonly WireSubNetOrdinalComparer Instance = new WireSubNetOrdinalComparer();

        public bool Equals((string Wire, string SubNet) x, (string Wire, string SubNet) y)
        {
            return string.Equals(x.Wire, y.Wire, StringComparison.OrdinalIgnoreCase)
                && string.Equals(x.SubNet, y.SubNet, StringComparison.OrdinalIgnoreCase);
        }

        public int GetHashCode((string Wire, string SubNet) obj)
        {
            var cmp = StringComparer.OrdinalIgnoreCase;
            return HashCode.Combine(cmp.GetHashCode(obj.Wire ?? ""), cmp.GetHashCode(obj.SubNet ?? ""));
        }
    }
}
