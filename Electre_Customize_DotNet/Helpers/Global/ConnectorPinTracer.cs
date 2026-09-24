using System;
using System.Collections.Generic;
using System.Linq;
using Electre_Customize_DotNet.Objects;

namespace Electre_Customize_DotNet.Helpers.Global
{
    /// <summary>
    /// Finds the pin on the other side of an EQU / break (_F/_M) connector, or the linked pins of a
    /// TBK (JM), SPL or TER component. The same rules serve the Power On and Continuity reports; each
    /// report passes the index over ITS OWN connection collection (whole project vs. selected looms).
    /// </summary>
    internal static class ConnectorPinTracer
    {
        public static ElectreObject TracePinofEQUConnector(ElectreTraversalIndex index, ElectreObject eleObj)
        {
            string lastConnectorName = eleObj.ConnectorName;
            // Iterate through keys in the connectorMap
            foreach (var key in EquConnectorMap.Pairs.Keys)
            {
                if (lastConnectorName.EndsWith(key, StringComparison.OrdinalIgnoreCase))
                {
                    string mappedValue = EquConnectorMap.Pairs[key];
                    int suffixIndex = lastConnectorName.Length - key.Length;
                    lastConnectorName = lastConnectorName.Substring(0, suffixIndex) + mappedValue;
                    break;
                }
            }

            var nextWireConnection = index.GetFirstByConnectorPin(lastConnectorName, eleObj.PinNumber);

            return nextWireConnection;
        }

        public static ElectreObject TracePinofBreakConnector(ElectreTraversalIndex index, ElectreObject eleObj)
        {
            string lastConnectorName = eleObj.ConnectorName;

            if (lastConnectorName.EndsWith("_F", StringComparison.OrdinalIgnoreCase))
            {
                lastConnectorName = lastConnectorName.Substring(0, lastConnectorName.Length - 2) + "_M";
            }
            else if (lastConnectorName.EndsWith("_M", StringComparison.OrdinalIgnoreCase))
            {
                lastConnectorName = lastConnectorName.Substring(0, lastConnectorName.Length - 2) + "_F";
            }

            var nextWireConnection = index.GetFirstByConnectorPin(lastConnectorName, eleObj.PinNumber);

            return nextWireConnection;
        }

        public static List<ElectreObject> TracePinOfJM(ElectreTraversalIndex index, ElectreObject eleObj)
        {
            var CoonnectedObjs = index.GetByConnectorTypeShunt(eleObj.ConnectorName, "TBK", eleObj.ShuntExt1)
                .Where(obj => obj.SubNet != eleObj.SubNet &&
                              !string.IsNullOrEmpty(obj.ShuntExt1))
                .ToList();

            return CoonnectedObjs;
        }

        public static List<ElectreObject> TracePinOfSPL(ElectreTraversalIndex index, ElectreObject eleObj)
        {
            var CoonnectedObjs = index.GetByConnectorTypeShunt(eleObj.ConnectorName, "SPL", eleObj.ShuntExt1)
                .Where(obj => obj.SubNet != eleObj.SubNet)
                .ToList();

            return CoonnectedObjs;
        }

        public static List<ElectreObject> TracePinOfTER(ElectreTraversalIndex index, ElectreObject eleObj)
        {
            var CoonnectedObjs = index.GetByConnectorTypeShunt(eleObj.ConnectorName, "TER", eleObj.ShuntExt1)
                                    .Where(obj => obj.SubNet != eleObj.SubNet &&
                                     !string.IsNullOrEmpty(obj.ShuntExt1))
                                     .ToList();

            return CoonnectedObjs;
        }
    }
}
