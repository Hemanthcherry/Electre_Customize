using Electre_Customize_DotNet.Objects;

namespace Electre_Customize_DotNet.Helpers.Global
{
    /// <summary>
    /// O(1)/O(degree) lookups that preserve original scan order and equality
    /// (ordinal Wire/SubNet/Pin, OrdinalIgnoreCase connector name, first match wins).
    /// </summary>
    internal sealed class ElectreTraceIndex
    {
        private static readonly List<ElectreObject> Empty = new List<ElectreObject>();

        private readonly Dictionary<(string Wire, string SubNet), List<ElectreObject>> _byWireSubNet;
        private readonly Dictionary<string, Dictionary<string, ElectreObject>> _byConnectorPin;
        private readonly Dictionary<(string Connector, string Type, string Shunt), List<ElectreObject>> _byShuntType;

        private ElectreTraceIndex(
            Dictionary<(string, string), List<ElectreObject>> byWireSubNet,
            Dictionary<string, Dictionary<string, ElectreObject>> byConnectorPin,
            Dictionary<(string, string, string), List<ElectreObject>> byShuntType)
        {
            _byWireSubNet = byWireSubNet;
            _byConnectorPin = byConnectorPin;
            _byShuntType = byShuntType;
        }

        public static ElectreTraceIndex Build(List<ElectreObject> source)
        {
            var byWire = new Dictionary<(string, string), List<ElectreObject>>();
            var byConnPin = new Dictionary<string, Dictionary<string, ElectreObject>>(StringComparer.OrdinalIgnoreCase);
            var byShunt = new Dictionary<(string, string, string), List<ElectreObject>>();

            if (source == null)
                return new ElectreTraceIndex(byWire, byConnPin, byShunt);

            for (int i = 0; i < source.Count; i++)
            {
                var obj = source[i];

                var wireKey = (obj.WireNumber, obj.SubNet);
                if (!byWire.TryGetValue(wireKey, out var wireList))
                {
                    wireList = new List<ElectreObject>(2);
                    byWire[wireKey] = wireList;
                }
                wireList.Add(obj);

                string connector = obj.ConnectorName ?? "";
                string pin = obj.PinNumber ?? "";
                if (!byConnPin.TryGetValue(connector, out var pins))
                {
                    pins = new Dictionary<string, ElectreObject>();
                    byConnPin[connector] = pins;
                }
                if (!pins.ContainsKey(pin))
                    pins[pin] = obj;

                string type = obj.ComponentType;
                if (type == "TBK" || type == "SPL" || type == "TER")
                {
                    var shuntKey = (obj.ConnectorName, type, obj.ShuntExt1);
                    if (!byShunt.TryGetValue(shuntKey, out var shuntList))
                    {
                        shuntList = new List<ElectreObject>(2);
                        byShunt[shuntKey] = shuntList;
                    }
                    shuntList.Add(obj);
                }
            }

            return new ElectreTraceIndex(byWire, byConnPin, byShunt);
        }

        public List<ElectreObject> ConnectedOnWire(string wireNumber, string subNet)
        {
            if (_byWireSubNet.TryGetValue((wireNumber, subNet), out var list))
                return list;
            return Empty;
        }

        public ElectreObject FindByConnectorPin(string connectorName, string pinNumber)
        {
            if (connectorName != null
                && _byConnectorPin.TryGetValue(connectorName, out var pins)
                && pins.TryGetValue(pinNumber ?? "", out var obj))
            {
                return obj;
            }
            return null;
        }

        public List<ElectreObject> ShuntPeers(ElectreObject eleObj, string componentType, bool requireNonEmptyShunt)
        {
            if (eleObj == null)
                return Empty;
            if (requireNonEmptyShunt && string.IsNullOrEmpty(eleObj.ShuntExt1))
                return Empty;

            if (!_byShuntType.TryGetValue((eleObj.ConnectorName, componentType, eleObj.ShuntExt1), out var list))
                return Empty;

            List<ElectreObject> result = null;
            for (int i = 0; i < list.Count; i++)
            {
                var obj = list[i];
                if (obj.SubNet != eleObj.SubNet)
                {
                    result ??= new List<ElectreObject>();
                    result.Add(obj);
                }
            }
            return result ?? Empty;
        }

        public static string MapEquConnectorName(string connectorName)
        {
            if (string.IsNullOrEmpty(connectorName))
                return connectorName;

            foreach (var key in EquConnectorMap.Pairs.Keys)
            {
                if (connectorName.EndsWith(key, StringComparison.OrdinalIgnoreCase))
                {
                    string mappedValue = EquConnectorMap.Pairs[key];
                    int suffixIndex = connectorName.Length - key.Length;
                    return connectorName.Substring(0, suffixIndex) + mappedValue;
                }
            }
            return connectorName;
        }

        public static string MapBreakConnectorName(string connectorName)
        {
            if (string.IsNullOrEmpty(connectorName))
                return connectorName;

            if (connectorName.EndsWith("_F", StringComparison.OrdinalIgnoreCase))
                return connectorName.Substring(0, connectorName.Length - 2) + "_M";
            if (connectorName.EndsWith("_M", StringComparison.OrdinalIgnoreCase))
                return connectorName.Substring(0, connectorName.Length - 2) + "_F";
            return connectorName;
        }
    }
}
