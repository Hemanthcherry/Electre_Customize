using Electre_Customize_DotNet.MainOperation;
using Electre_Customize_DotNet.Objects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Electre_Customize_DotNet.Helpers.Global;
using Electre_Customize_DotNet.Helpers.PowerOn;

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

        // The Trace*/FindGroundPin methods below always scan modMain.ElecCollection_All directly (not a
        // per-call parameter), so they share one lookup index that is rebuilt only when that list is
        // replaced by a new data load. See ElectreTraversalIndex for the exact semantics each lookup keeps.
        private static readonly TraversalIndexCache IndexCache = new TraversalIndexCache();

        internal static ElectreTraversalIndex GetIndex() => IndexCache.For(modMain.ElecCollection_All);

        /// <summary>
        /// Exposes the shared modMain.ElecCollection_All index to modExcel.TraceAndLogPath, which
        /// has its own copy of the same "objects sharing this wire+subnet" scan (mirrors
        /// source.Where(w => w.WireNumber == wireNumber &amp;&amp; w.SubNet == subNet)). Reuses the
        /// exact same cached index the Trace* methods below already use.
        /// </summary>
        public static List<ElectreObject> GetConnectedObjectsByWireSubNet(string wireNumber, string subNet)
        {
            return GetIndex().GetByWireSubNet(wireNumber, subNet);
        }

        public void PowerOnReportGeneration()
        {
            selectedPanelObject = PanelElectreObject(modMain.ElecCollection_All, _selectedPanellist);

            _modExcel.AppendToExcelPowerOn(selectedPanelObject);
        }

        private List<ElectreObject> PanelElectreObject(List<ElectreObject> elecollection, List<string> selectedPanel)
        {
            // elecollection is modMain.ElecCollection_All (project-wide scale). The original
            // re-lowered and re-scanned selectedPanel for every single element - lower it into a
            // HashSet once instead; e.Panel.ToLower() is left exactly as-is (still throws on a
            // null Panel, matching the original).
            var panelSet = new HashSet<string>(selectedPanel.Select(p => p.ToLower()));
            var collection = elecollection.Where(e => panelSet.Contains(e.Panel.ToLower())).ToList();
            return collection;
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
            // Same fix as PanelElectreObject above.
            var sheetSet = new HashSet<string>(selectedSheet.Select(p => p.ToLower()));
            var collection = elecollection.Where(e => sheetSet.Contains(e.SheetName.ToLower())).ToList();
            return collection;
        }
        #endregion

        #region //old FindGroundPin method commented on June 12
        /*  public static string FindGroundPin(List<ElectreObject> elecCollection, ElectreObject source)
          {
              // Step 1: Get all GNDs in the panel

              var gndObjects = elecCollection
              .Where(e => e.ComponentType.Contains("GROUND", StringComparison.OrdinalIgnoreCase) &&
                 e.Panel == source.Panel &&
                (e.ConnectorName.Split('_')[0].Equals(source.ConnectorName, StringComparison.OrdinalIgnoreCase) ||
                e.ConnectorName.Split('-')[0].Equals(source.ConnectorName, StringComparison.OrdinalIgnoreCase)))
              .ToList();


              if (gndObjects.Count == 0)
                  return ""; // No GND found

              var groundObj = gndObjects.First();

              string currentConnectorName = groundObj.ConnectorName;
              string currentPinNumber = groundObj.PinNumber;
              string currentWireNumber = groundObj.WireNumber;
              string tag = groundObj.Core_Part_Number;
              string subNet = groundObj.SubNet;

              // List<string> visitedWires = new List<string>();

              // List<string> visited = new List<string>();
              List<string> visitedConnections = new List<string>();

              // Step 2: Trace the continuity path
              while (!string.IsNullOrEmpty(currentWireNumber))
              {
                  // string visitKey = $"{currentConnectorName},{currentWireNumber}";

                  *//*  if (visited.Contains(visitKey))
                        break;*//* // Stop if we have already visited this connection

                  //  visited.Add(visitKey);
                  visitedConnections.Add($"{currentConnectorName},{currentPinNumber}");

                  *//*var connectedObjects = modMain.ElecCollection_All
                      .Where(w => w.WireNumber == currentWireNumber && !visitedConnections.Contains($"{w.ConnectorName},{w.PinNumber}") *//*!visitedConnectors.Contains(w.ConnectorName)*//*w.ConnectorName != lastConnectorName*//* && !string.IsNullOrEmpty(w.ConnectorName))
                      .ToList();*//*
                  var connectedObjects = modMain.ElecCollection_All
                      .Where(w => w.WireNumber + "," + w.Core_Part_Number + "," + w.SubNet == currentWireNumber + "," + tag + "," + subNet && !visitedConnections.Contains($"{w.ConnectorName},{w.PinNumber}")
                      *//*!visitedConnectors.Contains(w.ConnectorName)*//*w.ConnectorName != lastConnectorName*//* && !string.IsNullOrEmpty(w.ConnectorName))
                      .ToList();

                  if (connectedObjects.Count == 0)
                      break; // Stop if no further connections or loop detected


                  *//* if (connectedObjects.Count == 0 || visitedWires.Contains(currentWireNumber))
                       break; // Stop if no further connections or loop detected

                   visitedWires.Add(currentWireNumber);*//*

                  foreach (var connObj in connectedObjects)
                  {
                      if (connObj.ConnectorName == currentConnectorName && connObj.PinNumber == currentPinNumber)
                          continue; // Skip self

                      *//* if (visited.Contains($"{connObj.ConnectorName},{connObj.WireNumber}"))
                           continue;*//* // Prevent revisiting

                      currentConnectorName = connObj.ConnectorName;
                      currentPinNumber = connObj.PinNumber;

                      visitedConnections.Add($"{currentConnectorName},{currentPinNumber}");

                      switch (connObj.ComponentType)
                      {
                          case "EQU":
                              currentConnectorName = $"{connObj.ConnectorName}";
                              var nextConObj = ConnectorPinTracer.TracePinofEQUConnector(GetIndex(), connObj);
                              if(nextConObj != null)
                              {
                                  currentConnectorName = nextConObj.ConnectorName;
                              }
                             *//* if (currentConnectorName.EndsWith("M", StringComparison.OrdinalIgnoreCase )|| currentConnectorName.EndsWith("F", StringComparison.OrdinalIgnoreCase ))
                              {
                                  var brkConnectorName = TracePinofBreakConnectorGND(connObj);
                                  if (string.IsNullOrEmpty(brkConnectorName))
                                  {
                                      break;
                                  }
                                  currentConnectorName = brkConnectorName;
                                  break;
                              }*//*
                              break;
                          case "REL":
                          case "SWT":
                              *//* var ConnectorNameandPinSWT = TracePinOfSWT(connObj);
                               if (!string.IsNullOrEmpty(ConnectorNameandPinSWT))
                               {
                                   string[] parts = ConnectorNameandPinSWT.Split(',');
                                   currentConnectorName = parts[0];
                                   currentPinNumber = parts[1];
                               }*//*
                              break;
                          case "ERM":
                          case "DD":
                              var ConnectorNameandPinDD = TracePinOfDD(connObj);
                              if (!string.IsNullOrEmpty(ConnectorNameandPinDD))
                              {
                                  string[] parts = ConnectorNameandPinDD.Split(',');
                                  currentConnectorName = parts[0];
                                  currentPinNumber = parts[1];
                              }
                              break;
                          case "IND":
                          case "SCB":
                          case "TCB":
                          case "TER":
                              var ConnectorNameandPinTER = TracePinOfTERGND(connObj);
                              if (!string.IsNullOrEmpty(ConnectorNameandPinTER))
                              {
                                  string[] parts = ConnectorNameandPinTER.Split(',');
                                  currentConnectorName = parts[0];
                                  currentPinNumber = parts[1];
                              }
                              break;
                          case "FUS":
                          case "POT":
                          case "LMP":

                          case "ANT":
                          case "BUS":
                          case "MSW":
                              // currentConnectorName = $"%%{connObj.ConnectorName}#{connObj.PinNumber}%%";
                              currentConnectorName = $"{connObj.ConnectorName}";
                              break;
                          case "DIS":
                              var breakConnectorName = TracePinofBreakConnectorGND(connObj);
                              if (string.IsNullOrEmpty(breakConnectorName))
                              {
                                  break;
                              }
                              currentConnectorName = breakConnectorName;
                              break;

                          case "TBK":
                              var ConnectorNameandPinTBK = TracePinOfJMGND(connObj);
                              if (!string.IsNullOrEmpty(ConnectorNameandPinTBK))
                              {
                                  string[] parts = ConnectorNameandPinTBK.Split(',');
                                  currentConnectorName = parts[0];
                                  currentPinNumber = parts[1];
                              }
                              break;
                          case "SPL":
                              var ConnectorNameandPinSPL = TracePinOfSPLGND(connObj);
                              if (!string.IsNullOrEmpty(ConnectorNameandPinSPL))
                              {
                                  string[] parts = ConnectorNameandPinSPL.Split(',');
                                  currentConnectorName = parts[0];
                                  currentPinNumber = parts[1];
                              }
                              break;
                          default:
                              return currentPinNumber;//pass the for the empty value
                                                      // break;
                      }

                      var nextWireConnection = modMain.ElecCollection_All
                           .FirstOrDefault(w => string.Equals(w.ConnectorName, currentConnectorName, StringComparison.OrdinalIgnoreCase) && w.PinNumber == currentPinNumber);

                      if (nextWireConnection != null)
                      {
                          currentWireNumber = nextWireConnection.WireNumber;
                          currentConnectorName = nextWireConnection.ConnectorName;
                          currentPinNumber = nextWireConnection.PinNumber;
                          tag = nextWireConnection.Core_Part_Number;
                          subNet = nextWireConnection.SubNet;
                          break; // Move to the next wire
                      }
                  }
              }

              return $"{currentConnectorName},{currentPinNumber}"; // Return the final pin number for GND continuity
          }*/
        #endregion

        public static List<string> FindGroundPin(List<ElectreObject> elecCollection, ElectreObject source)
        {
            var gndObjects = elecCollection.Where(e =>
                                    (e.ComponentType.Contains("GROUND", StringComparison.OrdinalIgnoreCase) ||
                                     e.ComponentType.Contains("TER", StringComparison.OrdinalIgnoreCase) ||
                                     e.ComponentType.Contains("TBK", StringComparison.OrdinalIgnoreCase)) &&
                                    e.Panel == source.Panel &&
                                    e.ConnectorName.StartsWith($"{source.ConnectorName}_RTN", StringComparison.OrdinalIgnoreCase)
                                ).ToList();


            return TraceGroundPins(gndObjects);
        }

        /// <summary>
        /// Same result as <see cref="FindGroundPin(List{ElectreObject}, ElectreObject)"/>, but the ground candidates come from a
        /// per-panel index built once for the whole report instead of a full scan of the collection per circuit breaker.
        /// </summary>
        public static List<string> FindGroundPin(GroundCandidateIndex groundIndex, ElectreObject source)
        {
            return TraceGroundPins(groundIndex.CandidatesFor(source));
        }

        private static List<string> TraceGroundPins(List<ElectreObject> gndObjects)
        {
            List<string> groundConnectorAndPins = new List<string>();
           
            if (gndObjects.Count == 0)
                return new List<string>(); // No GND found

            foreach (var obj in gndObjects)
            {
                var visitedConnections = new HashSet<string>();
                TraceAndLogPathGND(obj, obj.WireNumber, obj.ConnectorName, obj.PinNumber, obj.SubNet, visitedConnections, ref groundConnectorAndPins);
            }

            return groundConnectorAndPins; // Return the final pin number for GND continuity
        }

        private static void TraceAndLogPathGND(ElectreObject source, string wireNumber, string connectorName, string pinNumber, string subNet, HashSet<string> visited, ref List<string> groundConnectorAndPins)
        {
            visited.Add($"{connectorName},{pinNumber}");

            var connectedObjects = GetIndex().GetByWireSubNet(wireNumber, subNet)
                .Where(w => !visited.Contains($"{w.ConnectorName},{w.PinNumber}"))
                .ToList();

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
                        var nextObjEQU = ConnectorPinTracer.TracePinofEQUConnector(GetIndex(), connObj);
                        if(nextObjEQU == null)
                        {
                            groundConnectorAndPins.Add($"{connObj.ConnectorName},{connObj.PinNumber}");
                            return; // If EQU is not connected to another connector, add it to the list and return
                        }
                        if (nextObjEQU != null)
                            TraceAndLogPathGND(source, nextObjEQU.WireNumber, nextObjEQU.ConnectorName, nextObjEQU.PinNumber, nextObjEQU.SubNet, visited, ref groundConnectorAndPins);
                        break;
                    case "DIS":
                        var nextObj = ConnectorPinTracer.TracePinofBreakConnector(GetIndex(), connObj);
                        if (nextObj == null)
                        {
                            groundConnectorAndPins.Add($"{connObj.ConnectorName},{connObj.PinNumber}");
                            return; // If DIS is not connected to another connector, add it to the list and return
                        }
                        if (nextObj != null)
                            TraceAndLogPathGND(source, nextObj.WireNumber, nextObj.ConnectorName, nextObj.PinNumber, nextObj.SubNet, visited, ref groundConnectorAndPins);
                        break;

                    case "TBK":
                        var jmList = ConnectorPinTracer.TracePinOfJM(GetIndex(), connObj);
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
                        var splList = ConnectorPinTracer.TracePinOfSPL(GetIndex(), connObj);
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
                        var terList = ConnectorPinTracer.TracePinOfTER(GetIndex(), connObj);
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

        /* public static string TracePinofBreakConnectorGND(ElectreObject eleObj)
       {
           string lastConnectorName = eleObj.ConnectorName;

           if (lastConnectorName.ToUpper().EndsWith("F"))
               lastConnectorName = lastConnectorName.Replace("F", "M");
           else if (lastConnectorName.ToUpper().EndsWith("M"))
               lastConnectorName = lastConnectorName.Replace("M", "F");

           var nextWireConnection = modMain.ElecCollection_All
                               .FirstOrDefault(w => string.Equals(w.ConnectorName, lastConnectorName, StringComparison.OrdinalIgnoreCase) && w.PinNumber == eleObj.PinNumber);

           if (nextWireConnection == null)
           {
               return "";
           }
           return lastConnectorName;
       }*/

        /*public static string TracePinOfJMGND(ElectreObject eleObj)
        {
            string TracePinOfJMResult = string.Empty;

            var collection = modMain.ElecCollection_All.Where(obj => obj.ConnectorName == eleObj.ConnectorName && obj.ComponentType == "TBK" && obj.WireNumber != eleObj.WireNumber && obj.ShuntExt1 == eleObj.ShuntExt1);

            foreach (var obj in collection)
            {
                TracePinOfJMResult = $"{obj.ConnectorName},{obj.PinNumber}";
                return TracePinOfJMResult;
            }
            return TracePinOfJMResult;
        }*/

        /* public static string TracePinOfSPLGND(ElectreObject eleObj)
         {
             string TracePinOfSPLResult = string.Empty;

             var CollectionObj = modMain.ElecCollection_All.Where(obj => obj.ConnectorName == eleObj.ConnectorName && obj.ComponentType == eleObj.ComponentType && obj.SubNet != eleObj.SubNet && obj.ComponentType == "SPL" && obj.ShuntExt1 == eleObj.ShuntExt1 && obj.WireNumber != eleObj.WireNumber);

             foreach (var obj in CollectionObj)
             {
                 TracePinOfSPLResult = $"{obj.ConnectorName},{obj.PinNumber}";
                 return TracePinOfSPLResult;
             }
             return TracePinOfSPLResult;
         }*/

        /*public static string TracePinOfSWT(ElectreObject eleObj)
        {
            string TracePinOfSWTResult = string.Empty;

            var CollectionObj = modMain.ElecCollection_All.Where(obj => obj.ConnectorName == eleObj.ConnectorName && obj.ComponentType == eleObj.ComponentType);

            foreach (var obj in CollectionObj)
            {
                if (obj.ComponentType == "SWT" && obj.PinNumber != eleObj.PinNumber)
                {
                    TracePinOfSWTResult = $"{obj.ConnectorName},{obj.PinNumber}";
                    return TracePinOfSWTResult;
                }
            }
            return TracePinOfSWTResult;
        }*/

        /* public static string TracePinOfTERGND(ElectreObject eleObj)
         {
             string TracePinOfTERResult = string.Empty;

             var CollectionObj = modMain.ElecCollection_All.Where(obj => obj.ConnectorName == eleObj.ConnectorName && obj.ComponentType == "TER" && obj.ShuntExt1 == eleObj.ShuntExt1 && obj.SubNet != eleObj.SubNet && !string.IsNullOrEmpty(obj.ShuntExt1));

             foreach (var obj in CollectionObj)
             {
                 TracePinOfTERResult = $"{obj.ConnectorName},{obj.PinNumber}";
                 return TracePinOfTERResult;
             }
             return TracePinOfTERResult;
         }*/

        /*public static string TracePinOfDD(ElectreObject eleObj)
        {
            string TracePinOfDDResult = string.Empty;

            var CollectionObj = modMain.ElecCollection_All.Where(obj => obj.ConnectorName == eleObj.ConnectorName && obj.ComponentType == eleObj.ComponentType);

            foreach (var obj in CollectionObj)
            {
                if (obj.ComponentType == "DD" && obj.PinNumber != eleObj.PinNumber)
                {
                    TracePinOfDDResult = $"{obj.ConnectorName},{obj.PinNumber}";
                    return TracePinOfDDResult;
                }
            }
            return TracePinOfDDResult;
        }*/

    }
}