using Electre_Customize_DotNet.Objects;
using System.Collections.Generic;

namespace Electre_Customize_DotNet.Helpers.Global
{
    /// <summary>
    /// Immutable, once-built lookup index over a snapshot of ElectreObjects, used by
    /// ContinuityWOBreakdown and PowerOnReport to replace repeated full-list
    /// .Where()/.FirstOrDefault() scans in the recursive trace code with dictionary lookups.
    /// It only changes how candidates are FOUND - callers still apply the same per-call filters
    /// (e.g. the visited-set check, the SubNet-inequality check) on the returned candidates, and
    /// each lookup's case sensitivity exactly mirrors the site it replaces (see method comments).
    /// Build one of these once per distinct source list, right after that list is finalized, and
    /// reuse it for every trace call against that same list.
    /// </summary>
    internal sealed class ElectreTraversalIndex
    {
        private static readonly List<ElectreObject> EmptyList = new List<ElectreObject>();

        // (WireNumber, SubNet) -> objects sharing that wire+subnet, in original list order.
        private readonly Dictionary<(string, string), List<ElectreObject>> _byWireSubNet
            = new Dictionary<(string, string), List<ElectreObject>>();

        // (ConnectorName upper-invariant, PinNumber) -> first matching object in original list
        // order. Case-insensitive connector match, exact-match pin - mirrors the
        // .FirstOrDefault(w => string.Equals(w.ConnectorName, x, StringComparison.OrdinalIgnoreCase)
        //                      && w.PinNumber == y) pattern used by the EQU/DIS lookups.
        private readonly Dictionary<(string, string), ElectreObject> _byConnectorPinFirst
            = new Dictionary<(string, string), ElectreObject>();

        // (ConnectorName exact, ComponentType, ShuntExt1) -> objects, in original list order.
        // Case-SENSITIVE connector match (obj.ConnectorName == x) - mirrors the JM/SPL/TER
        // lookups, which deliberately do not use OrdinalIgnoreCase unlike the EQU/DIS lookups.
        private readonly Dictionary<(string, string, string), List<ElectreObject>> _byConnectorTypeShunt
            = new Dictionary<(string, string, string), List<ElectreObject>>();

        public ElectreTraversalIndex(List<ElectreObject> source)
        {
            foreach (var obj in source)
            {
                var wireSubNetKey = (obj.WireNumber, obj.SubNet);
                if (!_byWireSubNet.TryGetValue(wireSubNetKey, out var wsList))
                {
                    wsList = new List<ElectreObject>();
                    _byWireSubNet[wireSubNetKey] = wsList;
                }
                wsList.Add(obj);

                var connectorPinKey = (CaseInsensitiveKey(obj.ConnectorName), obj.PinNumber);
                if (!_byConnectorPinFirst.ContainsKey(connectorPinKey))
                {
                    _byConnectorPinFirst[connectorPinKey] = obj; // first occurrence only, matches FirstOrDefault
                }

                var typeShuntKey = (obj.ConnectorName, obj.ComponentType, obj.ShuntExt1);
                if (!_byConnectorTypeShunt.TryGetValue(typeShuntKey, out var tsList))
                {
                    tsList = new List<ElectreObject>();
                    _byConnectorTypeShunt[typeShuntKey] = tsList;
                }
                tsList.Add(obj);
            }
        }

        private static string CaseInsensitiveKey(string s) => s == null ? null : s.ToUpperInvariant();

        /// <summary>
        /// Mirrors: source.Where(w => w.WireNumber == wireNumber &amp;&amp; w.SubNet == subNet).
        /// Returns the SAME shared list backing the index - callers must not mutate it (none do;
        /// they only read/iterate it and build their own filtered/derived lists from it).
        /// </summary>
        public List<ElectreObject> GetByWireSubNet(string wireNumber, string subNet)
        {
            return _byWireSubNet.TryGetValue((wireNumber, subNet), out var list) ? list : EmptyList;
        }

        /// <summary>
        /// Mirrors: source.FirstOrDefault(w =>
        ///     string.Equals(w.ConnectorName, connectorName, StringComparison.OrdinalIgnoreCase)
        ///     &amp;&amp; w.PinNumber == pinNumber)
        /// </summary>
        public ElectreObject GetFirstByConnectorPin(string connectorName, string pinNumber)
        {
            return _byConnectorPinFirst.TryGetValue((CaseInsensitiveKey(connectorName), pinNumber), out var obj)
                ? obj
                : null;
        }

        /// <summary>
        /// Mirrors: source.Where(obj =>
        ///     obj.ConnectorName == connectorName &amp;&amp; obj.ComponentType == componentType
        ///     &amp;&amp; obj.ShuntExt1 == shuntExt1)
        /// Same shared-list caveat as GetByWireSubNet.
        /// </summary>
        public List<ElectreObject> GetByConnectorTypeShunt(string connectorName, string componentType, string shuntExt1)
        {
            return _byConnectorTypeShunt.TryGetValue((connectorName, componentType, shuntExt1), out var list)
                ? list
                : EmptyList;
        }
    }
}
