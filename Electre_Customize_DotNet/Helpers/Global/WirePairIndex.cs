using Electre_Customize_DotNet.Objects;

namespace Electre_Customize_DotNet.Helpers.Global
{
    internal sealed class WirePairIndex
    {
        private readonly Dictionary<(string Wire, string SubNet), List<int>> _index;
        private readonly (string Wire, string SubNet)[] _keys;

        private WirePairIndex(
            Dictionary<(string Wire, string SubNet), List<int>> index,
            (string Wire, string SubNet)[] keys)
        {
            _index = index;
            _keys = keys;
        }

        public static WirePairIndex Build(List<ElectreObject> collection)
        {
            int n = collection.Count;
            var keys = new (string Wire, string SubNet)[n];
            var index = new Dictionary<(string, string), List<int>>(n, WireSubNetOrdinalComparer.Instance);
            for (int i = 0; i < n; i++)
            {
                var e = collection[i];
                if (string.IsNullOrEmpty(e.WireNumber) || string.Equals(e.ComponentType, "SDS", StringComparison.OrdinalIgnoreCase))
                    continue;

                var key = (e.WireNumber, e.SubNet ?? "");
                keys[i] = key;
                if (!index.TryGetValue(key, out var mates))
                {
                    mates = new List<int>(2);
                    index[key] = mates;
                }
                mates.Add(i);
            }
            return new WirePairIndex(index, keys);
        }

        public bool TryGetMates(int i, ElectreObject e, out List<int> mates)
        {
            (string Wire, string SubNet) key;
            if (_keys != null && i < _keys.Length && _keys[i].Wire != null)
                key = _keys[i];
            else
                key = (e.WireNumber ?? "", e.SubNet ?? "");

            return _index.TryGetValue(key, out mates);
        }
    }
}
