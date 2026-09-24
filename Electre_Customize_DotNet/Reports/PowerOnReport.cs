using Electre_Customize_DotNet.Helpers.Global;
using Electre_Customize_DotNet.Helpers.PowerOn;
using Electre_Customize_DotNet.MainOperation;
using Electre_Customize_DotNet.Objects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Electre_Customize_DotNet.Reports
{
    internal class PowerOnReport
    {
        private modExcel _modExcel;
        private static List<string> _selectedPanellist;

        public PowerOnReport(modExcel modExcel, List<string> selectedPanellist)
        {
            _modExcel = modExcel;
            _selectedPanellist = selectedPanellist;
        }

        public static List<ElectreObject> selectedPanelObject;
        public static List<ElectreObject> selectedSheetObject; // added for Panel Drawing schedules
        internal static ElectreTraceIndex AllIndex;
        private static GroundPinIndex _groundIndex;
        private static List<ElectreObject> _groundIndexSource;

        public void PowerOnReportGeneration()
        {
            selectedPanelObject = PanelElectreObject(modMain.ElecCollection_All, _selectedPanellist);

            _modExcel.AppendToExcelPowerOn(selectedPanelObject);
        }

        private List<ElectreObject> PanelElectreObject(List<ElectreObject> elecollection, List<string> selectedPanel)
        {
            return ElectreCollectionFilter.ByPanel(elecollection, selectedPanel);
        }

        #region Power on for Panel Drawing
        // added for Panel Drawing schedules
        public void PowerOnReportGeneration_PanelDWG()
        {
            selectedSheetObject = PanelElectreObject_PanelDWG(modMain.ElecCollection_All, _selectedPanellist);

            _modExcel.AppendToExcelPowerOn(selectedSheetObject);
        }

        // added for Panel Drawing schedules
        private List<ElectreObject> PanelElectreObject_PanelDWG(List<ElectreObject> elecollection, List<string> selectedSheet)
        {
            return ElectreCollectionFilter.BySheet(elecollection, selectedSheet);
        }
        #endregion

        public static List<string> FindGroundPin(List<ElectreObject> elecCollection, ElectreObject source)
        {
            AllIndex ??= ElectreTraceIndex.Build(modMain.ElecCollection_All);
            if (!ReferenceEquals(_groundIndexSource, elecCollection) || _groundIndex == null)
            {
                _groundIndex = GroundPinIndex.Build(elecCollection);
                _groundIndexSource = elecCollection;
            }

            var gndObjects = _groundIndex.FindForSource(source);

            List<string> groundConnectorAndPins = new List<string>();
            if (gndObjects.Count == 0)
                return groundConnectorAndPins;

            foreach (var obj in gndObjects)
            {
                var visitedConnections = new HashSet<string>();
                TraceAndLogPathGND(obj, obj.WireNumber, obj.ConnectorName, obj.PinNumber, obj.SubNet, visitedConnections, ref groundConnectorAndPins);
            }

            return groundConnectorAndPins;
        }

        private static void TraceAndLogPathGND(ElectreObject source, string wireNumber, string connectorName, string pinNumber, string subNet, HashSet<string> visited, ref List<string> groundConnectorAndPins)
        {
            visited.Add($"{connectorName},{pinNumber}");

            var wirePeers = AllIndex != null
                ? AllIndex.ConnectedOnWire(wireNumber, subNet)
                : modMain.ElecCollection_All;
            var connectedObjects = new List<ElectreObject>();
            for (int i = 0; i < wirePeers.Count; i++)
            {
                var w = wirePeers[i];
                if (w.WireNumber == wireNumber && w.SubNet == subNet
                    && !visited.Contains($"{w.ConnectorName},{w.PinNumber}"))
                {
                    connectedObjects.Add(w);
                }
            }

            if (connectedObjects.Count == 0)
            {
                groundConnectorAndPins.Add($"{connectorName},{pinNumber}");
                return;
            }

            foreach (var connObj in connectedObjects)
            {
                if (visited.Contains($"{connObj.ConnectorName},{connObj.PinNumber}"))
                    continue;

                switch (connObj.ComponentType)
                {
                    case "EQU":
                        var nextObjEQU = TracePinofEQUConnector(connObj);
                        if(nextObjEQU == null)
                        {
                            groundConnectorAndPins.Add($"{connObj.ConnectorName},{connObj.PinNumber}");
                            return; // If EQU is not connected to another connector, add it to the list and return
                        }
                        if (nextObjEQU != null)
                            TraceAndLogPathGND(source, nextObjEQU.WireNumber, nextObjEQU.ConnectorName, nextObjEQU.PinNumber, nextObjEQU.SubNet, visited, ref groundConnectorAndPins);
                        break;
                    case "DIS":
                        var nextObj = TracePinofBreakConnector(connObj);
                        if (nextObj == null)
                        {
                            groundConnectorAndPins.Add($"{connObj.ConnectorName},{connObj.PinNumber}");
                            return; // If DIS is not connected to another connector, add it to the list and return
                        }
                        if (nextObj != null)
                            TraceAndLogPathGND(source, nextObj.WireNumber, nextObj.ConnectorName, nextObj.PinNumber, nextObj.SubNet, visited, ref groundConnectorAndPins);
                        break;

                    case "TBK":
                        var jmList = TracePinOfJM(connObj);
                        if (jmList.Count == 0)
                        {
                            groundConnectorAndPins.Add($"{connObj.ConnectorName},{connObj.PinNumber}");
                            return; // If TBK is not connected to another connector, add it to the list and return
                        }
                        foreach (var jm in jmList)
                        {
                            TraceAndLogPathGND(source, jm.WireNumber, jm.ConnectorName, jm.PinNumber, jm.SubNet, visited, ref groundConnectorAndPins);
                        }
                        break;

                    case "SPL":
                        var splList = TracePinOfSPL(connObj);
                        if (splList.Count == 0)
                        {
                            groundConnectorAndPins.Add($"{connObj.ConnectorName},{connObj.PinNumber}");
                            return; // If SPL is not connected to another connector, add it to the list and return
                        }
                        foreach (var spl in splList)
                        {
                            TraceAndLogPathGND(source, spl.WireNumber, spl.ConnectorName, spl.PinNumber, spl.SubNet, visited, ref groundConnectorAndPins);
                        }
                        break;

                    case "TER":
                        var terList = TracePinOfTER(connObj);
                        if (terList.Count == 0)
                        {
                            groundConnectorAndPins.Add($"{connObj.ConnectorName},{connObj.PinNumber}");
                            return; // If TER is not connected to another connector, add it to the list and return
                        }
                        foreach (var ter in terList)
                        {
                            TraceAndLogPathGND(source, ter.WireNumber, ter.ConnectorName, ter.PinNumber, ter.SubNet, visited, ref groundConnectorAndPins);
                        }
                        break;

                    default:
                        groundConnectorAndPins.Add($"{connObj.ConnectorName},{connObj.PinNumber}");
                        return;
                       // TraceAndLogPathGND(source, connObj.WireNumber, connObj.ConnectorName, connObj.PinNumber, connObj.SubNet, visited, ref groundConnectorAndPins);
                       // break;
                }
            }
        }      

        public static ElectreObject TracePinofEQUConnector(ElectreObject eleObj)
        {
            return ConnectorTraceHops.TracePinofEQUConnector(eleObj, AllIndex, modMain.ElecCollection_All);
        }

        public static ElectreObject TracePinofBreakConnector(ElectreObject eleObj)
        {
            return ConnectorTraceHops.TracePinofBreakConnector(eleObj, AllIndex, modMain.ElecCollection_All);
        }

        public static List<ElectreObject> TracePinOfJM(ElectreObject eleObj)
        {
            return ConnectorTraceHops.TracePinOfJM(eleObj, AllIndex, modMain.ElecCollection_All);
        }

        public static List<ElectreObject> TracePinOfSPL(ElectreObject eleObj)
        {
            return ConnectorTraceHops.TracePinOfSPL(eleObj, AllIndex, modMain.ElecCollection_All);
        }

        public static List<ElectreObject> TracePinOfTER(ElectreObject eleObj)
        {
            return ConnectorTraceHops.TracePinOfTER(eleObj, AllIndex, modMain.ElecCollection_All);
        }

    }
}