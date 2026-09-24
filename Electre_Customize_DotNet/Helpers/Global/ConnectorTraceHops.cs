using Electre_Customize_DotNet.Objects;

namespace Electre_Customize_DotNet.Helpers.Global
{
    /// <summary>
    /// Shared EQU / DIS / TBK / SPL / TER hop lookups.
    /// Continuity and Power On use the same mapping and shunt rules;
    /// they pass different indexes (loom vs full collection).
    /// </summary>
    internal static class ConnectorTraceHops
    {
        public static ElectreObject? TracePinofEQUConnector(ElectreObject eleObj, ElectreTraceIndex? index, List<ElectreObject>? fallback)
        {
            string lastConnectorName = ElectreTraceIndex.MapEquConnectorName(eleObj.ConnectorName);
            if (index != null)
                return index.FindByConnectorPin(lastConnectorName, eleObj.PinNumber);

            return fallback?.FirstOrDefault(w =>
                string.Equals(w.ConnectorName, lastConnectorName, StringComparison.OrdinalIgnoreCase)
                && w.PinNumber == eleObj.PinNumber);
        }

        public static ElectreObject? TracePinofBreakConnector(ElectreObject eleObj, ElectreTraceIndex? index, List<ElectreObject>? fallback)
        {
            string lastConnectorName = ElectreTraceIndex.MapBreakConnectorName(eleObj.ConnectorName);
            if (index != null)
                return index.FindByConnectorPin(lastConnectorName, eleObj.PinNumber);

            return fallback?.FirstOrDefault(w =>
                string.Equals(w.ConnectorName, lastConnectorName, StringComparison.OrdinalIgnoreCase)
                && w.PinNumber == eleObj.PinNumber);
        }

        public static List<ElectreObject> TracePinOfJM(ElectreObject eleObj, ElectreTraceIndex? index, List<ElectreObject>? fallback)
        {
            if (index != null)
                return index.ShuntPeers(eleObj, "TBK", requireNonEmptyShunt: true);

            if (fallback == null)
                return new List<ElectreObject>();

            return fallback
                .Where(obj => obj.ConnectorName == eleObj.ConnectorName &&
                              obj.ComponentType == "TBK" &&
                              obj.SubNet != eleObj.SubNet &&
                              obj.ShuntExt1 == eleObj.ShuntExt1 &&
                              !string.IsNullOrEmpty(obj.ShuntExt1))
                .ToList();
        }

        public static List<ElectreObject> TracePinOfSPL(ElectreObject eleObj, ElectreTraceIndex? index, List<ElectreObject>? fallback)
        {
            if (index != null)
                return index.ShuntPeers(eleObj, "SPL", requireNonEmptyShunt: false);

            if (fallback == null)
                return new List<ElectreObject>();

            return fallback
                .Where(obj => obj.ConnectorName == eleObj.ConnectorName &&
                              obj.ComponentType == "SPL" &&
                              obj.SubNet != eleObj.SubNet &&
                              obj.ShuntExt1 == eleObj.ShuntExt1)
                .ToList();
        }

        public static List<ElectreObject> TracePinOfTER(ElectreObject eleObj, ElectreTraceIndex? index, List<ElectreObject>? fallback)
        {
            if (index != null)
                return index.ShuntPeers(eleObj, "TER", requireNonEmptyShunt: true);

            if (fallback == null)
                return new List<ElectreObject>();

            return fallback
                .Where(obj => obj.ConnectorName == eleObj.ConnectorName &&
                              obj.ComponentType == "TER" &&
                              obj.ShuntExt1 == eleObj.ShuntExt1 &&
                              obj.SubNet != eleObj.SubNet &&
                              !string.IsNullOrEmpty(obj.ShuntExt1))
                .ToList();
        }
    }
}
