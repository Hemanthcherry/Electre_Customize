using Electre_Customize_DotNet.Logs;
using Electre_Customize_DotNet.MainOperation;
using Electre_Customize_DotNet.Objects;
using System.Configuration;
using System.Windows.Forms;

namespace Electre_Customize_DotNet.Reports
{
    internal class ContinuityWOBreakdown
    {
        private modExcel _modExcel;
        private List<string> _selectedLoomlist;

        public ContinuityWOBreakdown(modExcel modExcel, List<string> selectedLoomlist)
        {
            _modExcel = modExcel;
            _selectedLoomlist = selectedLoomlist;
        }

        public static List<ElectreObject> selectedLoomObject;

        public void ContinuityReportGeneration(string ContinuityFolder, string ContinuityCompFolder, List<ElectreObject> filteredElecCollection )
        {
            ContinuityWOBreakDown(filteredElecCollection);

            string reportName = "Continuity_WO_Breakdown";

            string[,] arrFTcwobDist = removeRedundants(modMain.arrFTcwob);           

            _modExcel.GenerateHALReportFormat_ContinuityList(arrFTcwobDist, ContinuityFolder, reportName);

            // Sheet Continuity Components starts 
            //string workbookPath = _modExcel.CreateNewWorkbook("Continuity_Components", "Continuity_Components", ConfigurationManager.AppSettings["ContinuityCompFolder"]);
            string workbookPath = _modExcel.CreateNewWorkbook("Continuity_Components", "Continuity_Components", ContinuityCompFolder);

            if (!_modExcel.CreateListOfContinuityComponentsReportHeader(workbookPath))
            {
                Logging.Error("Failed to create the wireList Report Header");
            }

            var listOfContinuityComponents = modMain.arrListOfContinuityComponents;

            Dictionary<string, List<string>> connectorsAndPartNames = modMain.connectorsAndPartNames;

            Dictionary<string, List<string>> matchedComponents = new Dictionary<string, List<string>>();

            foreach (var continuityComponent in listOfContinuityComponents)
            {
                if (connectorsAndPartNames.ContainsKey(continuityComponent))
                {
                    var uniquePartNames = connectorsAndPartNames[continuityComponent].Distinct().ToList();
                    matchedComponents.Add(continuityComponent, uniquePartNames);
                }
            }
            _modExcel.AppendContinuityComponentsandPartNumber(matchedComponents);
        }

        void ContinuityWOBreakDown(List<ElectreObject> filteredElecCollection)
        {
            try
            {
                // clear the previous data if exists
                modMain.arrListOfContinuityComponents.Clear();
                modMain.destinationConnectors.Clear();
                // clearing the arrFTcwob object
                if (modMain.arrFTcwob != null && modMain.arrFTcwob.Cast<string>().Any(value => !string.IsNullOrEmpty(value)))
                {
                    // If arrFTcwob contains data, clear it
                    for (int i = 0; i < modMain.arrFTcwob.GetLength(0); i++)
                    {
                        for (int j = 0; j < modMain.arrFTcwob.GetLength(1); j++)
                        {
                            modMain.arrFTcwob[i, j] = null; // Clear data
                        }
                    }
                    Console.WriteLine("arrFTcwob data cleared.");
                }
                try
                {
                    //filters the selected loom object from ElecCollection based on the selected loom list
                    selectedLoomObject = loomElectreObject(filteredElecCollection, _selectedLoomlist);

                    // Sort the selectedLoomObject based on the specified criteria
                    selectedLoomObject = selectedLoomObject.OrderByDescending(obj => obj.ComponentType == "EQU") // "EQU" first
                                                            .ThenBy(obj =>
                                                                modMain.connectorMap.Keys.Any(suffix =>
                                                                    obj.ConnectorName.EndsWith($"_{suffix}", StringComparison.OrdinalIgnoreCase)
                                                                ) ? 1 : 0 // if it matches a known connector suffix → 1 (comes later), else → 0 (comes first)
                                                            )
                                                            .ThenBy(obj => obj.ConnectorName, StringComparer.OrdinalIgnoreCase)
                                                            .ThenBy(obj => ExtractNumericPrefix(obj.PinNumber))
                                                            .ThenBy(obj => obj.PinNumber, StringComparer.OrdinalIgnoreCase)
                                                            .ToList();

                    // Filter out objects with No Megger is not null, then add it into ElecCollectionforNoMegger 
                    modMain.ElecCollectionforNoMegger = selectedLoomObject.Where(e => !string.IsNullOrEmpty(e.NoMegger)).ToList();

                    // Filter out objects with No Megger is null, then add it into selectedLoomObject for continuity termination
                    selectedLoomObject = selectedLoomObject.Where(e => string.IsNullOrEmpty(e.NoMegger)).ToList();
                }
                catch (Exception ex)
                {
                    Logging.Error("ContinuityWOBreakDown : Error while filtering or sorting selected loom objects: " + ex.Message);
                    return;
                }
                
                // Components where both left and right side exists
                var matchingBreakConnectors = processMatchingConnectors();
                int rowNumber = 0;

                foreach (var source in selectedLoomObject)
                {
                    var visitedConnections = new HashSet<string>();

                    if (matchingBreakConnectors.Contains($"{source.ConnectorName},{source.PinNumber}") ||
                        modMain.destinationConnectors.Contains($"{source.ConnectorName},{source.PinNumber}"))
                    {
                        continue;
                    }

                    TraceAndAppendData(source, source.WireNumber, source.ConnectorName, source.PinNumber, source.SubNet, visitedConnections, ref rowNumber);
                }

            }
            catch (Exception ex)
            {
                Logging.Error("Error in ContinuityWOBreakDown: " + ex.Message);
            }
        }

        public HashSet<string> processMatchingConnectors()
        {
            HashSet<string> matchingBreakConnectors = new HashSet<string>();
            try
            {
                // Step 1: (Break Connectors DIS) Group connectors by their base name (excluding _M or _F)
                var groupedBreakConnectors = selectedLoomObject
                                             .Where(obj => obj.ConnectorName.EndsWith("_M", StringComparison.OrdinalIgnoreCase) ||
                                             obj.ConnectorName.EndsWith("_F", StringComparison.OrdinalIgnoreCase))
                                             .GroupBy(obj => obj.ConnectorName.Substring(0, obj.ConnectorName.Length - 2))
                                             .ToList();

                foreach (var group in groupedBreakConnectors)
                {
                    var maleConnectors = group
                        .Where(obj => obj.ConnectorName.EndsWith("_M", StringComparison.OrdinalIgnoreCase))
                        .ToList();

                    var femaleConnectors = group
                        .Where(obj => obj.ConnectorName.EndsWith("_F", StringComparison.OrdinalIgnoreCase))
                        .ToList();

                    if (maleConnectors.Any() && femaleConnectors.Any())
                    {
                        foreach (var conn in maleConnectors.Concat(femaleConnectors))
                        {
                            if (!string.IsNullOrEmpty(conn.PinNumber))
                            {
                                matchingBreakConnectors.Add($"{conn.ConnectorName},{conn.PinNumber}");
                            }
                        }
                    }
                }

                // Step 2 : (EQU Connectors) Group all connectors by their base name (before the suffix) with matching suffixes left and right
                var groupedEQUConnectors = selectedLoomObject
                    .Select(obj =>
                    {
                        // Find matching suffix
                        var matchingSuffix = modMain.connectorMap.Keys
                            .FirstOrDefault(suffix => obj.ConnectorName.EndsWith($"{suffix}", StringComparison.OrdinalIgnoreCase));

                        if (matchingSuffix != null)
                        {
                            // Return object with base name (without suffix)
                            return new
                            {
                                BaseName = obj.ConnectorName.Substring(0, obj.ConnectorName.Length - matchingSuffix.Length - 1), // subtract '_suffix'
                                Electre = obj
                            };
                        }
                        return null;
                    })
                    .Where(x => x != null)
                    .GroupBy(x => x.BaseName)
                    .ToList();

                // Male connector suffixes: J1 to J24
                var maleSuffixes = new List<string>
                {
                    "J1", "J2", "J3", "J4", "J5", "J6", "J7", "J8",
                    "J9", "J10", "J11", "J12", "J13", "J14", "J15", "J16",
                    "J17", "J18", "J19", "J20", "J21", "J22", "J23", "J24"
                };

                // Female connector suffixes: a to z, excluding i and o
                var femaleSuffixes = new List<string>
                {
                    "a", "b", "c", "d", "e", "f", "g", "h",
                    "j", "k", "l", "m", "n", "p", "q", "r",
                    "s", "t", "u", "v", "w", "x", "y", "z"
                };

                // Process each group and find valid male/female pairs
                foreach (var group in groupedEQUConnectors)
                {
                    var maleConnectors = group
                        .Where(x => maleSuffixes.Any(suffix => x.Electre.ConnectorName.EndsWith($"{suffix}", StringComparison.OrdinalIgnoreCase)))
                        .Select(x => x.Electre)
                        .ToList();

                    var femaleConnectors = group
                        .Where(x => femaleSuffixes.Any(suffix => x.Electre.ConnectorName.EndsWith($"{suffix}", StringComparison.OrdinalIgnoreCase)))
                        .Select(x => x.Electre)
                        .ToList();

                    if (maleConnectors.Any() && femaleConnectors.Any())
                    {
                        foreach (var conn in maleConnectors.Concat(femaleConnectors))
                        {
                            if (!string.IsNullOrEmpty(conn.PinNumber))
                            {
                                matchingBreakConnectors.Add($"{conn.ConnectorName},{conn.PinNumber}");
                            }
                        }
                    }
                }

                // Step 3: Group all TER, TBK and SPL components by their Shunt
                string[] ComponentTypes = { "TER", "TBK", "SPL" };
                var groupedByShuntAndConnector = selectedLoomObject
                                            .Where(obj => ComponentTypes.Contains(obj.ComponentType) && !string.IsNullOrEmpty(obj.Shunt))
                                            .GroupBy(obj => new { obj.Shunt, obj.ConnectorName });

                foreach (var group in groupedByShuntAndConnector)
                {
                    var componentsInGroup = group.ToList();

                    if (componentsInGroup.Count >= 2)
                    {
                        foreach (var component in componentsInGroup)
                        {
                            //if (!string.IsNullOrEmpty(component.PinNumber))
                            //{
                                string connectorandPin = $"{component.ConnectorName},{component.PinNumber}";
                                if (!matchingBreakConnectors.Contains(connectorandPin))
                                {
                                    matchingBreakConnectors.Add(connectorandPin);
                                }
                           // }
                        }
                    }
                }

                return matchingBreakConnectors;
            }           
            catch (Exception ex)
            {
                Logging.Error("ContinuityWOBreakDown: Error while processing matchingBreakConnectors: " + ex.Message);
                return matchingBreakConnectors;
            }

        }

        private void TraceAndAppendData(ElectreObject source, string wireNumber, string connectorName, string pinNumber, string subNet, HashSet<string> visited, ref int rowNumber)
        {
            try
            {
                visited.Add($"{connectorName},{pinNumber}");

                var connectedObjects = selectedLoomObject
                    .Where(w => w.WireNumber == wireNumber && w.SubNet == subNet && !visited.Contains($"{w.ConnectorName},{w.PinNumber}"))
                    .ToList();

                if (connectedObjects.Count == 0)
                {
                    if(source.ConnectorName == connectorName && source.PinNumber == pinNumber) // if source and destination are same then return
                    {
                        return;
                    }
                    AppendContinuityDatato2dArray(source, connectorName, pinNumber, ref rowNumber);
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
                            if (nextObjEQU == null)
                            {
                                AppendContinuityDatato2dArray(source, connObj.ConnectorName, connObj.PinNumber, ref rowNumber);
                                return;
                            }
                            if (nextObjEQU != null)
                                TraceAndAppendData(source, nextObjEQU.WireNumber, nextObjEQU.ConnectorName, nextObjEQU.PinNumber, nextObjEQU.SubNet, visited, ref rowNumber);
                            break;
                        case "DIS":
                            var nextObj = TracePinofBreakConnector(connObj);
                            if (nextObj == null)
                            {
                                AppendContinuityDatato2dArray(source, connObj.ConnectorName, connObj.PinNumber, ref rowNumber);
                                return;
                            }
                            if (nextObj != null)
                                TraceAndAppendData(source, nextObj.WireNumber, nextObj.ConnectorName, nextObj.PinNumber, nextObj.SubNet, visited, ref rowNumber);
                            break;

                        case "TBK":
                            var jmList = TracePinOfJM(connObj);
                            if (jmList.Count == 0)
                            {
                                AppendContinuityDatato2dArray(source, connObj.ConnectorName, connObj.PinNumber, ref rowNumber);
                                return;
                            }
                            foreach (var jm in jmList)
                            {
                                TraceAndAppendData(source, jm.WireNumber, jm.ConnectorName, jm.PinNumber, jm.SubNet, visited, ref rowNumber);
                            }
                            break;

                        case "SPL":
                            var splList = TracePinOfSPL(connObj);
                            if (splList.Count == 0)
                            {
                                AppendContinuityDatato2dArray(source, connObj.ConnectorName, connObj.PinNumber, ref rowNumber);
                                return;
                            }
                            foreach (var spl in splList)
                            {
                                TraceAndAppendData(source, spl.WireNumber, spl.ConnectorName, spl.PinNumber, spl.SubNet, visited, ref rowNumber);
                            }
                            break;

                        case "TER":
                            var terList = TracePinOfTER(connObj);
                            if (terList.Count == 0)
                            {
                                AppendContinuityDatato2dArray(source, connObj.ConnectorName, connObj.PinNumber, ref rowNumber);
                                return;
                            }
                            foreach (var ter in terList)
                            {
                                TraceAndAppendData(source, ter.WireNumber, ter.ConnectorName, ter.PinNumber, ter.SubNet, visited, ref rowNumber);
                            }
                            break;

                        default:
                            AppendContinuityDatato2dArray(source, connObj.ConnectorName, connObj.PinNumber, ref rowNumber);
                            //TraceAndLogPath(source, connObj.WireNumber, connObj.ConnectorName, connObj.PinNumber, connObj.SubNet, visited, negPins, ref parameterPINs);
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                Logging.Error($"TraceAndAppendData Error: Source={source?.ConnectorName},{source?.PinNumber} and Dest={connectorName},{pinNumber} - {ex.Message}");
            }
        }

        public static ElectreObject TracePinofEQUConnector(ElectreObject eleObj)
        {
            string lastConnectorName = eleObj.ConnectorName;
            // Iterate through keys in the connectorMap
            foreach (var key in modMain.connectorMap.Keys)
            {
                if (lastConnectorName.EndsWith(key, StringComparison.OrdinalIgnoreCase))
                {
                    string mappedValue = modMain.connectorMap[key];
                    int suffixIndex = lastConnectorName.Length - key.Length;
                    lastConnectorName = lastConnectorName.Substring(0, suffixIndex) + mappedValue;
                    break;
                }
            }

            var nextWireConnection = selectedLoomObject
                                .FirstOrDefault(w => string.Equals(w.ConnectorName, lastConnectorName, StringComparison.OrdinalIgnoreCase) && w.PinNumber == eleObj.PinNumber);

            return nextWireConnection;
        }

        public static ElectreObject TracePinofBreakConnector(ElectreObject eleObj)
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

            var nextWireConnection = selectedLoomObject
                                .FirstOrDefault(w => string.Equals(w.ConnectorName, lastConnectorName, StringComparison.OrdinalIgnoreCase) && w.PinNumber == eleObj.PinNumber);

            return nextWireConnection;
        }

        public static List<ElectreObject> TracePinOfJM(ElectreObject eleObj)
        {
            var CoonnectedObjs = selectedLoomObject
                .Where(obj => obj.ConnectorName == eleObj.ConnectorName &&
                              obj.ComponentType == "TBK" &&
                              obj.SubNet != eleObj.SubNet &&
                              obj.ShuntExt1 == eleObj.ShuntExt1 &&
                              !string.IsNullOrEmpty(obj.ShuntExt1))
                .ToList();

            return CoonnectedObjs;
        }

        public static List<ElectreObject> TracePinOfSPL(ElectreObject eleObj)
        {
            var CoonnectedObjs = selectedLoomObject
                .Where(obj => obj.ConnectorName == eleObj.ConnectorName &&
                              obj.ComponentType == "SPL" &&
                              obj.SubNet != eleObj.SubNet &&
                              obj.ShuntExt1 == eleObj.ShuntExt1)
                .ToList();

            return CoonnectedObjs;
        }

        public static List<ElectreObject> TracePinOfTER(ElectreObject eleObj)
        {
            var CoonnectedObjs = selectedLoomObject
                                    .Where(obj => obj.ConnectorName == eleObj.ConnectorName &&
                                     obj.ComponentType == "TER" &&
                                     obj.ShuntExt1 == eleObj.ShuntExt1 &&
                                     obj.SubNet != eleObj.SubNet &&
                                     !string.IsNullOrEmpty(obj.ShuntExt1))
                                     .ToList();

            return CoonnectedObjs;
        }

        void AppendContinuityDatato2dArray(ElectreObject source, string destConnector, string destPin, ref int rowNumber)
        {
            try
            {
                if (rowNumber < modMain.arrFTcwob.GetLength(0))
                {
                    modMain.arrFTcwob[rowNumber, 0] = !string.IsNullOrEmpty(source.ConnectorName) ? source.ConnectorName : "NULL";
                    modMain.arrFTcwob[rowNumber, 1] = !string.IsNullOrEmpty(source.PinNumber) ? source.PinNumber : "NULL";
                    modMain.arrFTcwob[rowNumber, 5] = !string.IsNullOrEmpty(source.SheetName) ? source.SheetName : "NULL";
                    modMain.arrFTcwob[rowNumber, 6] = modMain.layerAssign(source.Layer); // Assuming this method handles null internally
                    modMain.arrFTcwob[rowNumber, 2] = !string.IsNullOrEmpty(destConnector) ? destConnector : "NULL";
                    modMain.arrFTcwob[rowNumber, 3] = !string.IsNullOrEmpty(destPin) ? destPin : "NULL";

                    string core = null;

                    // Check if Core_Part_Number is an integer between 1 and 18
                    if (int.TryParse(source.Core_Part_Number, out int coreNum) && coreNum >= 1 && coreNum <= 18)
                    {
                        core = source.Core_Part_Number;
                    }
                    // Check if the CableType is in the specified string sCableType_O
                    if (modMain.sCableType_O.Contains("_" + source.CableType + "_"))
                    {
                        if (source.Core_Part_Number == "" && source.CableType != "X")
                        {
                            modMain.arrFTcwob[rowNumber, 4] = $"{source.WireNumber}/{(!string.IsNullOrEmpty(source.Gauge) && source.Gauge.Length > 1 ? source.Gauge.Substring(1) : "")}";
                        }
                        else
                        {
                            // modMain.arrFTcwob[rowNumber, 4] = source.WireNumber + "/" + source.Gauge.Substring(1) + "/" + source.Core_Part_Number;
                            modMain.arrFTcwob[rowNumber, 4] = source.WireNumber + "/" + source.Gauge.Substring(1) + "/" + core;
                        }
                    }
                    else if (modMain.sCableType_1.Contains($"_{source.CableType}_"))
                    {
                        modMain.arrFTcwob[rowNumber, 4] = $"{source.WireNumber}/{core}";
                        // modMain.arrFTcwob[rowNumber, 4] = $"{source.WireNumber}/{source.Core_Part_Number}";//{E1.Tag7} add in the between wirenumber and tag3 if required
                    }
                    else if (modMain.sCableType_2.Contains($"_{source.CableType}_"))
                    {
                        modMain.arrFTcwob[rowNumber, 4] = $"{source.WireNumber}/{core}";
                        // modMain.arrFTcwob[rowNumber, 4] = $"{source.WireNumber}/{source.Core_Part_Number}";//{E1.Tag7} add in the between wirenumber and tag3 if required
                    }
                    else
                    {
                        if (source.Core_Part_Number == "" && source.CableType != "X")
                        {
                            modMain.arrFTcwob[rowNumber, 4] = $"{source.WireNumber}/{(!string.IsNullOrEmpty(source.Gauge) && source.Gauge.Length > 1 ? source.Gauge.Substring(1) : "")}";
                        }
                        else
                        {
                            modMain.arrFTcwob[rowNumber, 4] = source.WireNumber + "/" + source.Gauge.Substring(1) + "/" + core;
                            //modMain.arrFTcwob[rowNumber, 4] = source.WireNumber + "/" + source.Gauge.Substring(1) + "/" + source.Core_Part_Number;
                        }
                    }                
                    rowNumber++;
                }
                else
                {
                    Logging.Error($"AppendContinuityDatato2dArray: rowNumber {rowNumber} exceeds arrFTcwob row limit.");                    
                }

                if (!string.IsNullOrEmpty(source.ConnectorName))
                {
                    modMain.SearchAndAppend(source.ConnectorName, ref modMain.arrListOfContinuityComponents);
                }
                if (!string.IsNullOrEmpty(destConnector))
                {
                    modMain.SearchAndAppend(destConnector, ref modMain.arrListOfContinuityComponents);
                    modMain.SearchAndAppend($"{destConnector},{destPin}", ref modMain.destinationConnectors); // adding destination connector and pin to avoid reverse continuity
                }
            }
            catch(Exception ex)
            {
                Logging.Error($"AppendContinuityDatato2dArray Error => {ex.Message}");
            }            
        }

        // Helper function to extract numeric part from PinNumber for proper sorting
        private static int ExtractNumericPrefix(string pinNumber)
        {
            string numericPart = new string(pinNumber.TakeWhile(char.IsDigit).ToArray());
            return int.TryParse(numericPart, out int result) ? result : int.MaxValue; // Non-numeric pins go last
        }

        private List<ElectreObject> loomElectreObject(List<ElectreObject> elecollection, List<string> selectedLoom)
        {
            return elecollection.Where(e => selectedLoom.Contains(e.BundleName)).ToList();
        }

        public bool IfExist(string istr, List<string> iarr)
        {
            for (int i = 0; i < iarr.Count; i++)
            {
                if (iarr[i] == istr)
                {
                    return true;
                }
            }
            return false;
        }

        // Helper methods
        bool FoundInList(string item, List<string> list)
        {
            return list.Contains(item);
        }

        int PinSortKey(string pin)
        {
            // If numeric, return integer value
            if (int.TryParse(pin, out int num))
                return num;

            // If alphabetical, convert to base-26 integer (like Excel columns: A=1, B=2, ..., AA=27)
            int value = 0;
            foreach (char c in pin.ToUpper())
            {
                if (char.IsLetter(c))
                {
                    value = value * 26 + (c - 'A' + 1);
                }
                else
                {
                    // Fallback for unexpected format
                    return int.MaxValue;
                }
            }
            return value;
        }

        // Helper functions and variables

        // Function to check if the ComponentType is valid
        bool IsValidComponentType(string componentType)
        {
            var validTypes = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
            "EQU", "REL", "SWT", "ERM", "IND", "SCB", "TCB", "TER",
            "FUS", "POT", "LMP", "ANT", "BUS", "MSW","DIS"
            };
            return validTypes.Contains(componentType);
        }

        // removing duplicate connections and destinations for continuity
        public string[,] removeRedundants(string[,] arrFTcwob)
        {
           // arrFTcwob = Sort2DArrayByFirstColumn(arrFTcwob);
            int totalRows = arrFTcwob.GetLength(0);
            string[,] arrFTcwobDist = new string[totalRows, 7];

            HashSet<string> rededuntCheck = new HashSet<string>();
            HashSet<string> duplicateDestTracker = new HashSet<string>();
            Dictionary<string, int> destPinCounts = new Dictionary<string, int>();

            int arrFTcwobDistCount = 0;

            // Count destination pairs
            for (int i = 0; i < totalRows; i++)
            {
                string destKey = $"{arrFTcwob[i, 2]}#{arrFTcwob[i, 3]}";
                if (!destPinCounts.TryAdd(destKey, 1))
                    destPinCounts[destKey]++;
            }

            for (int i = 0; i < totalRows; i++)
            {
                string src = $"{arrFTcwob[i, 0]}#{arrFTcwob[i, 1]}";
                string dest = $"{arrFTcwob[i, 2]}#{arrFTcwob[i, 3]}";

                string presentLine = $"{src}##{dest}";
                string connectLine = $"{dest}##{src}";

                if (destPinCounts[dest] > 1 && duplicateDestTracker.Contains(dest))
                    continue;

                if (!rededuntCheck.Contains(presentLine) && !rededuntCheck.Contains(connectLine) && presentLine != "####" && connectLine !="####")
                {
                    for (int col = 0; col < 7; col++)
                        arrFTcwobDist[arrFTcwobDistCount, col] = arrFTcwob[i, col];

                    arrFTcwobDistCount++;
                    rededuntCheck.Add(presentLine);

                    if (destPinCounts[dest] > 1)
                        duplicateDestTracker.Add(dest);
                }
            }

            // Trim the result to only filled rows
            string[,] result = new string[arrFTcwobDistCount, 7];
            for (int i = 0; i < arrFTcwobDistCount; i++)
                for (int j = 0; j < 7; j++)
                    result[i, j] = arrFTcwobDist[i, j];

            return result; 
        }

        // removing duplicate connections and destinations for Megger
        public string[,] removeRedundantsMegger(string[,] arrFTcwob)
        {
           // arrFTcwob = Sort2DArrayByFirstColumn(arrFTcwob);
            int totalRows = arrFTcwob.GetLength(0);
            string[,] arrFTcwobDist = new string[totalRows, 4];

            HashSet<string> rededuntCheck = new HashSet<string>();
            HashSet<string> duplicateDestTracker = new HashSet<string>();
            Dictionary<string, int> destPinCounts = new Dictionary<string, int>();

            int arrFTcwobDistCount = 0;

            // Count destination pairs
            for (int i = 0; i < totalRows; i++)
            {
                string destKey = $"{arrFTcwob[i, 2]}#{arrFTcwob[i, 3]}";
                if (!destPinCounts.TryAdd(destKey, 1))
                    destPinCounts[destKey]++;
            }

            for (int i = 0; i < totalRows; i++)
            {
                string src = $"{arrFTcwob[i, 0]}#{arrFTcwob[i, 1]}";
                string dest = $"{arrFTcwob[i, 2]}#{arrFTcwob[i, 3]}";

                string presentLine = $"{src}##{dest}";
                string connectLine = $"{dest}##{src}";

                if (destPinCounts[dest] > 1 && duplicateDestTracker.Contains(dest))
                    continue;

                if (!rededuntCheck.Contains(presentLine) && !rededuntCheck.Contains(connectLine))
                {
                    for (int col = 0; col < 4; col++)
                        arrFTcwobDist[arrFTcwobDistCount, col] = arrFTcwob[i, col];

                    arrFTcwobDistCount++;
                    rededuntCheck.Add(presentLine);

                    if (destPinCounts[dest] > 1)
                        duplicateDestTracker.Add(dest);
                }
            }

            return arrFTcwobDist;
        }

        public string[,] Sort2DArrayByFirstColumn(string[,] input)
        {
            int rowCount = input.GetLength(0);
            int colCount = input.GetLength(1);

            // Step 1: Convert 2D array to a list of string arrays (rows)
            var rows = new List<string[]>(rowCount);
            for (int i = 0; i < rowCount; i++)
            {
                string[] row = new string[colCount];
                for (int j = 0; j < colCount; j++)
                {
                    row[j] = input[i, j];
                }
                rows.Add(row);
            }

            // step 2: Sort list by first column (index 0)
            rows.Sort((a, b) => string.Compare(a[0], b[0], StringComparison.OrdinalIgnoreCase));

            // Step 3: Convert back to a 2D array
            string[,] sortedArray = new string[rowCount, colCount];
            for (int i = 0; i < rowCount; i++)
            {
                for (int j = 0; j < colCount; j++)
                {
                    sortedArray[i, j] = rows[i][j];
                }
            }

            return sortedArray;
        }

        public void ConnectionComponentsforMegger(List<ElectreObject> filteredElecCollection)
        {
            ContinuityWOBreakDown(filteredElecCollection);
        }

        #region Commented on July 7th code of Continuity before implementing Latest code like Power On
        // Commented on July 7th before implementing Power On like code
        /*void ContinuityWOBreakDown()
        {
            try
            {
                // clear the previous data if exists
                modMain.arrListOfContinuityComponents.Clear();
                modMain.destinationConnectors.Clear();
                // clearing the arrFTcwob object
                if (modMain.arrFTcwob != null && modMain.arrFTcwob.Cast<string>().Any(value => !string.IsNullOrEmpty(value)))
                {
                    // If arrFTcwob contains data, clear it
                    for (int i = 0; i < modMain.arrFTcwob.GetLength(0); i++)
                    {
                        for (int j = 0; j < modMain.arrFTcwob.GetLength(1); j++)
                        {
                            modMain.arrFTcwob[i, j] = null; // Clear data
                        }
                    }
                    Console.WriteLine("arrFTcwob data cleared.");
                }
                try
                {
                    //filters the selected loom object from ElecCollection based on the selected loom list
                    selectedLoomObject = loomElctreObject(modMain.ElecCollection, _selectedLoomlist);

                    // Sort the selectedLoomObject based on the specified criteria
                    selectedLoomObject = selectedLoomObject.OrderByDescending(obj => obj.ComponentType == "EQU") // "EQU" first
                                                            .ThenBy(obj =>
                                                                modMain.connectorMap.Keys.Any(suffix =>
                                                                    obj.ConnectorName.EndsWith($"_{suffix}", StringComparison.OrdinalIgnoreCase)
                                                                ) ? 1 : 0 // if it matches a known connector suffix → 1 (comes later), else → 0 (comes first)
                                                            )
                                                            .ThenBy(obj => obj.ConnectorName, StringComparer.OrdinalIgnoreCase)
                                                            .ThenBy(obj => ExtractNumericPrefix(obj.PinNumber))
                                                            .ThenBy(obj => obj.PinNumber, StringComparer.OrdinalIgnoreCase)
                                                            .ToList();

                    // Filter out objects with No Megger is not null, then add it into ElecCollectionforNoMegger 
                    modMain.ElecCollectionforNoMegger = selectedLoomObject.Where(e => !string.IsNullOrEmpty(e.NoMegger)).ToList();

                    // Filter out objects with No Megger is null, then add it into selectedLoomObject for continuity termination
                    selectedLoomObject = selectedLoomObject.Where(e => string.IsNullOrEmpty(e.NoMegger)).ToList();
                }
                catch (Exception ex)
                {
                    Logging.Error("ContinuityWOBreakDown : Error while filtering or sorting selected loom objects: " + ex.Message);
                    return;
                }

                // Variable declarations
                ElectreObject E1;
                ElectreObject n1;
                int rowNumber = 0; // Assuming rowNumber is defined and initialized
                rowNumber += 1; // Added this line due to header overwriting the initial line. Didn't want to change initial variable

                // Created List for storing Break Connectors, Connectors and Terminal Blocks if both left and right side exists in the selectedLoomObject to avoid as Source
                // List<string> matchingBreakConnectors = new List<string>();
                HashSet<string> matchingBreakConnectors = new HashSet<string>();
                try
                {
                    // Group connectors by their base name (excluding _M or _F)
                    var groupedBreakConnectors = selectedLoomObject
            .Where(obj => obj.ConnectorName.EndsWith("_M", StringComparison.OrdinalIgnoreCase) ||
                          obj.ConnectorName.EndsWith("_F", StringComparison.OrdinalIgnoreCase))
            .GroupBy(obj => obj.ConnectorName.Substring(0, obj.ConnectorName.Length - 2))
            .ToList();

                    foreach (var group in groupedBreakConnectors)
                    {
                        var maleConnectors = group
                            .Where(obj => obj.ConnectorName.EndsWith("_M", StringComparison.OrdinalIgnoreCase))
                            .ToList();

                        var femaleConnectors = group
                            .Where(obj => obj.ConnectorName.EndsWith("_F", StringComparison.OrdinalIgnoreCase))
                            .ToList();

                        if (maleConnectors.Any() && femaleConnectors.Any())
                        {
                            foreach (var conn in maleConnectors.Concat(femaleConnectors))
                            {
                                if (!string.IsNullOrEmpty(conn.PinNumber))
                                {
                                    matchingBreakConnectors.Add($"{conn.ConnectorName},{conn.PinNumber}");
                                }
                            }
                        }
                    }

                    // Group all connectors by their base name (before the suffix) with matching suffixes left and right
                    var groupedEQUConnectors = selectedLoomObject
                        .Select(obj =>
                        {
                            // Find matching suffix
                            var matchingSuffix = modMain.connectorMap.Keys
                                .FirstOrDefault(suffix => obj.ConnectorName.EndsWith($"{suffix}", StringComparison.OrdinalIgnoreCase));

                            if (matchingSuffix != null)
                            {
                                // Return object with base name (without suffix)
                                return new
                                {
                                    BaseName = obj.ConnectorName.Substring(0, obj.ConnectorName.Length - matchingSuffix.Length - 1), // subtract '_suffix'
                                    Electre = obj
                                };
                            }
                            return null;
                        })
                        .Where(x => x != null)
                        .GroupBy(x => x.BaseName)
                        .ToList();

                    // Male connector suffixes: J1 to J24
                    var maleSuffixes = new List<string>
                {
                    "J1", "J2", "J3", "J4", "J5", "J6", "J7", "J8",
                    "J9", "J10", "J11", "J12", "J13", "J14", "J15", "J16",
                    "J17", "J18", "J19", "J20", "J21", "J22", "J23", "J24"
                };

                    // Female connector suffixes: a to z, excluding i and o
                    var femaleSuffixes = new List<string>
                {
                    "a", "b", "c", "d", "e", "f", "g", "h",
                    "j", "k", "l", "m", "n", "p", "q", "r",
                    "s", "t", "u", "v", "w", "x", "y", "z"
                };

                    // Process each group and find valid male/female pairs
                    foreach (var group in groupedEQUConnectors)
                    {
                        var maleConnectors = group
                            .Where(x => maleSuffixes.Any(suffix => x.Electre.ConnectorName.EndsWith($"{suffix}", StringComparison.OrdinalIgnoreCase)))
                            .Select(x => x.Electre)
                            .ToList();

                        var femaleConnectors = group
                            .Where(x => femaleSuffixes.Any(suffix => x.Electre.ConnectorName.EndsWith($"{suffix}", StringComparison.OrdinalIgnoreCase)))
                            .Select(x => x.Electre)
                            .ToList();

                        if (maleConnectors.Any() && femaleConnectors.Any())
                        {
                            foreach (var conn in maleConnectors.Concat(femaleConnectors))
                            {
                                if (!string.IsNullOrEmpty(conn.PinNumber))
                                {
                                    matchingBreakConnectors.Add($"{conn.ConnectorName},{conn.PinNumber}");
                                }
                            }
                        }
                    }

                    // Group all terminal block entries by their Shunt
                    var terminalGroups = selectedLoomObject
                                                .Where(obj => obj.ComponentType == "TER" && !string.IsNullOrEmpty(obj.Shunt))
                                                .GroupBy(obj => new { obj.Shunt, obj.ConnectorName });

                    foreach (var group in terminalGroups)
                    {
                        var terminals = group.ToList();

                        if (terminals.Count >= 2)
                        {
                            foreach (var terminal in terminals)
                            {
                                if (!string.IsNullOrEmpty(terminal.PinNumber))
                                {
                                    string connectorWithPin = $"{terminal.ConnectorName},{terminal.PinNumber}";
                                    if (!matchingBreakConnectors.Contains(connectorWithPin))
                                    {
                                        matchingBreakConnectors.Add(connectorWithPin);
                                    }
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Logging.Error("ContinuityWOBreakDown: Error while processing connectors and terminals in matchingBreakConnectors: " + ex.Message);
                    return;
                }

                try
                {
                    // Assuming ElecCollection is a List<ElectreObject>
                    for (int i = 0; i < selectedLoomObject.Count; i++) // Adjusted for zero-based indexing
                    {
                        E1 = selectedLoomObject[i];
                        if (matchingBreakConnectors.Contains($"{E1.ConnectorName},{E1.PinNumber}") ||
                           modMain.destinationConnectors.Contains($"{E1.ConnectorName},{E1.PinNumber}"))
                        {
                            continue;
                        }

                        if (!string.IsNullOrEmpty(E1.WireNumber) && IsValidComponentType(E1.ComponentType))
                        {
                            for (int J = 0; J < selectedLoomObject.Count; J++)
                            {
                                n1 = selectedLoomObject[J];

                                if (J != i) // To omit this line
                                {
                                    if (string.Equals(n1.WireNumber, E1.WireNumber, StringComparison.OrdinalIgnoreCase) &&
                                        string.Equals(n1.SubNet, E1.SubNet, StringComparison.OrdinalIgnoreCase) &&
                                        n1.ComponentType != "SDS")
                                    {
                                        string t1;
                                        switch (n1.ComponentType)
                                        {
                                            case "EQU":
                                                t1 = $"%%{n1.ConnectorName}#{n1.PinNumber}%%";
                                                if (modMain.connectorMap.Keys.Any(k => n1.ConnectorName.EndsWith(k, StringComparison.OrdinalIgnoreCase)))
                                                {
                                                    t1 = $"%%{TracePinOfBreakConnector(n1.ConnectorName, n1.PinNumber, i, n1.ComponentType)}";
                                                }
                                                PrintConnectedLRU(E1, t1, ref rowNumber);
                                                break;
                                            case "REL":
                                            case "SWT":
                                            case "ERM":
                                            case "IND":
                                            case "SCB":
                                            case "TCB":
                                            case "FUS":
                                            case "POT":
                                            case "LMP":
                                            case "ANT":
                                            case "BUS":
                                            case "MSW":
                                            case "SNR":
                                                t1 = $"%%{n1.ConnectorName}#{n1.PinNumber}%%";
                                                PrintConnectedLRU(E1, t1, ref rowNumber);
                                                break;

                                            case "DIS":
                                                t1 = $"%%{TracePinOfBreakConnector(n1.ConnectorName, n1.PinNumber, i, n1.ComponentType)}";
                                                PrintConnectedLRU(E1, t1, ref rowNumber);
                                                break;

                                            case "TBK":
                                            case "TER":
                                            case "SPL":
                                                t1 = $"%%{TracePinOfJM(n1)}";
                                                PrintConnectedLRU(E1, t1, ref rowNumber);
                                                break;
                                            default:
                                                // t1 = "%%" + "" + "#" + "" + "%%";//pass the for the empty value
                                                t1 = $"%%{n1.ConnectorName}#{n1.PinNumber}%%";
                                                PrintConnectedLRU(E1, t1, ref rowNumber);
                                                break;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Logging.Error("ContinuityWOBreakDown: Error while processing continuity objects: " + ex.Message);
                    return;
                }
            }
            catch (Exception ex)
            {
                Logging.Error("Error in ContinuityWOBreakDown: " + ex.Message);
            }
        }*/

        // Commented on July 7th 2025 before implementing Power On like code
        /* string TracePinOfBreakConnector(string iConnectorName, string iPinNumber, int iElecCollectionIndexSearched, string componentType)
         {
             try
             {
                 // Variables declaration
                 string BreakConnectorName = string.Empty;
                 string GivenSide = string.Empty;
                 string? MatchingSide = string.Empty; // M-Movable, F-Fixed
                 ElectreObject b1;
                 ElectreObject C1;
                 string TracePinOfBreakConnectorResult = "";
                 // Other side of iConnectorName
                 string ConnectingConnectorName = string.Empty;

                 if (componentType.ToUpper() == "EQU")
                 {
                     foreach (var kvp in modMain.connectorMap)
                     {
                         if (iConnectorName.EndsWith(kvp.Key, StringComparison.OrdinalIgnoreCase))
                         {
                             // GivenSide = kvp.Key;
                             // MatchingSide = kvp.Value;
                             int suffixIndex = iConnectorName.Length - kvp.Key.Length;
                             ConnectingConnectorName = iConnectorName.Substring(0, suffixIndex) + kvp.Value;
                             // BreakConnectorName = iConnectorName.Substring(0, iConnectorName.Length - kvp.Key.Length);
                             break;
                         }
                     }
                 }
                 else if (componentType.ToUpper() == "DIS")
                 {
                     // BreakConnectorName = iConnectorName.Substring(0, iConnectorName.Length - 1);
                     // Get the last character of iConnectorName

                     if (iConnectorName.EndsWith("_F", StringComparison.OrdinalIgnoreCase))
                     {
                         ConnectingConnectorName = iConnectorName.Substring(0, iConnectorName.Length - 2) + "_M";
                     }
                     else if (iConnectorName.EndsWith("_M", StringComparison.OrdinalIgnoreCase))
                     {
                         ConnectingConnectorName = iConnectorName.Substring(0, iConnectorName.Length - 2) + "_F";
                     }
                     else
                     {
                         ConnectingConnectorName = iConnectorName; // or handle as invalid case
                     }
                     //GivenSide = iConnectorName.Substring(iConnectorName.Length - 2);

                     //// Determine MatchingSide based on GivenSide
                     //if (GivenSide.ToUpper() == "_F")
                     //{
                     //    MatchingSide = "_M";
                     //}
                     //else if (GivenSide.ToUpper() == "_M")
                     //{
                     //    MatchingSide = "_F";
                     //}
                     //else
                     //{
                     //    MatchingSide = "";
                     //}
                     //ConnectingConnectorName = iConnectorName.Replace(GivenSide, MatchingSide);
                 }

                 //string ConnectingConnectorName = iConnectorName.Replace(GivenSide, MatchingSide);
                 // string ConnectingConnectorName = BreakConnectorName + MatchingSide;

                 // Loop through ElecCollection
                 for (int b = 0; b < selectedLoomObject.Count; b++)
                 {
                     b1 = selectedLoomObject[b];
                     if (b1.ConnectorName.Equals(ConnectingConnectorName, StringComparison.OrdinalIgnoreCase) && b1.PinNumber == iPinNumber)
                     {
                         // Inner loop
                         for (int c = 0; c < selectedLoomObject.Count; c++)
                         {
                             C1 = selectedLoomObject[c];

                             // Check conditions
                             if (!string.IsNullOrEmpty(C1.WireNumber) &&
                                 C1.WireNumber == b1.WireNumber &&
                                 C1.SubNet == b1.SubNet &&
                                 C1.ConnectorName != b1.ConnectorName)
                             {
                                 switch (C1.ComponentType)
                                 {
                                     case "EQU":
                                         // TracePinOfBreakConnectorResult = C1.ConnectorName + "#" + C1.PinNumber + "%%" + TracePinOfBreakConnectorResult;
                                         TracePinOfBreakConnectorResult = C1.ConnectorName + "#" + C1.PinNumber + "%%";
                                         c = selectedLoomObject.Count;

                                         if (modMain.connectorMap.Keys.Any(k => C1.ConnectorName.EndsWith(k, StringComparison.OrdinalIgnoreCase)))
                                         {
                                             //TracePinOfBreakConnectorResult = TracePinOfBreakConnector(C1.ConnectorName, C1.PinNumber, c, C1.ComponentType) + TracePinOfBreakConnectorResult;
                                             TracePinOfBreakConnectorResult = TracePinOfBreakConnector(C1.ConnectorName, C1.PinNumber, c, C1.ComponentType);
                                         }
                                         break;
                                     case "REL":
                                     case "SWT":
                                     case "ERM":
                                     case "IND":
                                     case "SCB":
                                     case "TCB":
                                     // case "TER":
                                     case "FUS":
                                     case "POT":
                                     case "LMP":
                                     case "ANT":
                                     case "BUS":
                                     case "MSW":
                                         TracePinOfBreakConnectorResult = C1.ConnectorName + "#" + C1.PinNumber + "%%" + TracePinOfBreakConnectorResult;
                                         // Exit the inner loop
                                         c = selectedLoomObject.Count;
                                         break;

                                     case "DIS":
                                         // Recursive call
                                         TracePinOfBreakConnectorResult = TracePinOfBreakConnector(C1.ConnectorName, C1.PinNumber, c, C1.ComponentType) + TracePinOfBreakConnectorResult;
                                         break;

                                     case "TBK":
                                     case "TER":
                                     case "SPL":
                                         // Call TracePinOfJM function (needs to be implemented)
                                         TracePinOfBreakConnectorResult = TracePinOfJM(C1) + TracePinOfBreakConnectorResult;
                                         break;
                                 }
                             }
                         }
                     }
                 }

                 // If result is empty, assign default value
                 string[] parts = TracePinOfBreakConnectorResult.Split('#');
                 if (parts.Length >= 2)
                 {
                     string connectorName = parts[0];
                     string[] pinParts = parts[1].Split(new[] { "%%" }, StringSplitOptions.None);
                     string pinNumber = pinParts.Length > 0 ? pinParts[0] : string.Empty;

                     if (!string.IsNullOrEmpty(connectorName) && !string.IsNullOrEmpty(pinNumber) &&
                         !modMain.destinationConnectors.Contains($"{connectorName},{pinNumber}"))
                     {
                         modMain.destinationConnectors.Add($"{connectorName},{pinNumber}");
                     }
                     Logging.Info("TraceOfPinBC: " + TracePinOfBreakConnectorResult);
                 }
                 else
                 {
                     Logging.Warning("TracePinOfBreakConnectorResult is not in expected format: " + TracePinOfBreakConnectorResult);
                 }                

                 return TracePinOfBreakConnectorResult;
             }
             catch (Exception ex)
             {
                 Logging.Error($"Error in TracePinOfBreakConnector (Connector: {iConnectorName}, Pin: {iPinNumber}): {ex.Message}");
                 return iConnectorName + "#" + iPinNumber + "%%"; // Fallback safe value
             }
         }*/

        // Commented on July 7th 2025 before implementing Power On like code
        /* string TracePinOfJM(ElectreObject iJMpinInstance, string iInterJM_FromWire = "")
         {
             try
             {
                 ElectreObject F1;
                 ElectreObject E1;
                 ElectreObject G1;
                 string InterJM_FromWire = iInterJM_FromWire;
                 List<string> arrJMpinWithSameLink = new List<string>();
                 List<string> arrSPLwireWithSameLink = new List<string>();
                 List<string> arrTerwireWithSameLink = new List<string>();
                 bool Process1 = false;
                 string TracePinOfJMResult = "";

                 F1 = iJMpinInstance;

                 for (int E = 0; E < selectedLoomObject.Count; E++)
                 {
                     E1 = selectedLoomObject[E];

                     // Conditions for Splice
                     if (E1.ComponentType == "SPL" && E1.ConnectorName == F1.ConnectorName && E1.SubNet != F1.SubNet && E1.ShuntExt1 == F1.ShuntExt1 && E1.WireNumber != F1.WireNumber)
                     {
                         if (!FoundInList(E1.WireNumber, arrSPLwireWithSameLink))
                         {
                             modMain.SearchAndAppend(E1.WireNumber, ref arrSPLwireWithSameLink);
                             Process1 = true;
                         }
                     }
                     // Conditions for Terminal Block (TER)
                     else if (E1.ComponentType == "TER" && E1.ConnectorName == F1.ConnectorName && E1.WireNumber != F1.WireNumber && E1.ShuntExt1 == F1.ShuntExt1 && !string.IsNullOrEmpty(E1.ShuntExt1))
                     {
                         if (!FoundInList(E1.PinNumber, arrTerwireWithSameLink))
                         {
                             modMain.SearchAndAppend(E1.PinNumber, ref arrTerwireWithSameLink);
                             Process1 = true;
                         }
                     }
                     // Conditions for Junction Module (JM)
                     else if (E1.ComponentType == "TBK" && E1.ConnectorName == F1.ConnectorName && E1.PinNumber != F1.PinNumber && E1.ShuntExt1 == F1.ShuntExt1)
                     {
                         if (!FoundInList(E1.PinNumber, arrJMpinWithSameLink))
                         {
                             modMain.SearchAndAppend(E1.PinNumber, ref arrJMpinWithSameLink);
                             Process1 = true;
                         }
                     }

                     if (Process1)
                     {
                         for (int G = 0; G < selectedLoomObject.Count; G++)
                         {
                             G1 = selectedLoomObject[G];
                             if (!string.IsNullOrEmpty(G1.WireNumber) &&
                                 G1.WireNumber == E1.WireNumber &&
                                 G1.SubNet == E1.SubNet &&
                                 G1.ConnectorName != E1.ConnectorName &&
                                 G1.WireNumber != InterJM_FromWire)
                             {
                                 switch (G1.ComponentType)
                                 {
                                     case "EQU":
                                         TracePinOfJMResult = G1.ConnectorName + "#" + G1.PinNumber + "%%" + TracePinOfJMResult;
                                         G = selectedLoomObject.Count;
                                         if (modMain.connectorMap.Keys.Any(k => G1.ConnectorName.EndsWith(k, StringComparison.OrdinalIgnoreCase)))
                                         {
                                             TracePinOfJMResult = TracePinOfBreakConnector(G1.ConnectorName, G1.PinNumber, G, G1.ComponentType) + TracePinOfJMResult;
                                         }
                                         break;
                                     case "REL":
                                     case "SWT":
                                     case "ERM":
                                     case "IND":
                                     case "SCB":
                                     case "TCB":
                                     //case "TER":
                                     case "FUS":
                                     case "POT":
                                     case "LMP":
                                     case "ANT":
                                     case "BUS":
                                     case "MSW":
                                     case "SNR":
                                         TracePinOfJMResult = G1.ConnectorName + "#" + G1.PinNumber + "%%" + TracePinOfJMResult;
                                         // Exit the loop
                                         G = selectedLoomObject.Count;
                                         break;

                                     case "DIS":
                                         TracePinOfJMResult = TracePinOfBreakConnector(G1.ConnectorName, G1.PinNumber, G, G1.ComponentType) + TracePinOfJMResult;
                                         break;

                                     case "TBK":
                                     case "TER":
                                     case "SPL":
                                         string a1 = TracePinOfJMResult;
                                         string b1 = TracePinOfJM(G1, E1.WireNumber);
                                         string t1 = b1 + a1;
                                         TracePinOfJMResult = t1;
                                         break;
                                 }
                             }
                         }
                         Process1 = false;
                     }
                 }

                 if (string.IsNullOrEmpty(TracePinOfJMResult))
                 {
                     TracePinOfJMResult = F1.ConnectorName + "#" + F1.PinNumber + "%%";
                 }
                 string[] parts = TracePinOfJMResult.Split('#');
                 string connectorName = parts.Length > 0 ? parts[0] : "";
                 string pinNumber = parts.Length > 1 ? parts[1].Split(new[] { "%%" }, StringSplitOptions.None)[0] : "";

                 if (!string.IsNullOrEmpty(connectorName) && !string.IsNullOrEmpty(pinNumber) &&
                     !modMain.destinationConnectors.Contains($"{connectorName},{pinNumber}"))
                 {
                     modMain.destinationConnectors.Add($"{connectorName},{pinNumber}");
                 }
                 Logging.Info("TraceOfPinJM: " + TracePinOfJMResult);
                 //Console.WriteLine("TraceOfPinJM: " + TracePinOfJMResult);
                 return TracePinOfJMResult;
             }
             catch (Exception ex)
             {
                 Logging.Error("TracePinOfJM failed: " + ex.Message);
                 return "";
             }
         }*/

        // Commented on July 7th 2025 before implementing Power On like code
        /*void PrintConnectedLRU(ElectreObject iElecObj, string istr1, ref int rowNumber)
        {
            try
            {
                if (string.IsNullOrEmpty(istr1))
                {
                    Logging.Warning("PrintConnectedLRU: Input string is null or empty.");
                    return;
                }
                string[] CollectiveValues = istr1.Split(new[] { "%%" }, StringSplitOptions.None);

                if (CollectiveValues.Length < 2)
                {
                    Logging.Warning($"PrintConnectedLRU: Insufficient collective values: {istr1}");
                    return;
                }

                for (int T = 1; T <= CollectiveValues.Length - 2; T++)
                {
                    string tracedPart = CollectiveValues[T];
                    if (string.IsNullOrEmpty(tracedPart))
                        continue;

                    string[] TracedValue = tracedPart.Split('#');
                    if (TracedValue.Length < 2)
                    {
                        Logging.Warning($"PrintConnectedLRU: Invalid traced value format: '{tracedPart}'");
                        continue;
                    }
                    string ConnectedEQU = TracedValue[0] ?? "";
                    string ConnectedPin = TracedValue[1] ?? "";

                    //string str9 = iElecObj.ConnectorName + iElecObj.PinNumber;
                    string str9 = (iElecObj.ConnectorName ?? "") + (iElecObj.PinNumber ?? "");
                    string str8 = ConnectedEQU + ConnectedPin;

                    if (!string.Equals(str8, str9, StringComparison.OrdinalIgnoreCase))
                    {
                        // Defensive: Check array bounds before assignment
                        if (rowNumber < modMain.arrFTcwob.GetLength(0))
                        {
                            modMain.arrFTcwob[rowNumber, 0] = iElecObj.ConnectorName ?? "";
                            modMain.arrFTcwob[rowNumber, 1] = iElecObj.PinNumber ?? "";
                            modMain.arrFTcwob[rowNumber, 5] = iElecObj.SheetName ?? "";
                            modMain.arrFTcwob[rowNumber, 6] = modMain.layerAssign(iElecObj.Layer);

                            // Check if the CableType is in the specified string sCableType_O
                            if (modMain.sCableType_O.Contains("_" + iElecObj.CableType + "_"))
                            {
                                if (iElecObj.Core_Part_Number == "" && iElecObj.CableType != "X")
                                {
                                    modMain.arrFTcwob[rowNumber, 4] = $"{iElecObj.WireNumber}/{(!string.IsNullOrEmpty(iElecObj.Gauge) && iElecObj.Gauge.Length > 1 ? iElecObj.Gauge.Substring(1) : "")}";
                                }
                                else
                                {
                                    modMain.arrFTcwob[rowNumber, 4] = iElecObj.WireNumber + "/" + iElecObj.Gauge.Substring(1) + "/" + iElecObj.Core_Part_Number;
                                }
                            }
                            else if (modMain.sCableType_1.Contains($"_{iElecObj.CableType}_"))
                            {
                                modMain.arrFTcwob[rowNumber, 4] = $"{iElecObj.WireNumber}/{iElecObj.Core_Part_Number}";//{E1.Tag7} add in the between wirenumber and tag3 if required
                            }
                            else if (modMain.sCableType_2.Contains($"_{iElecObj.CableType}_"))
                            {
                                modMain.arrFTcwob[rowNumber, 4] = $"{iElecObj.WireNumber}/{iElecObj.Core_Part_Number}";//{E1.Tag7} add in the between wirenumber and tag3 if required
                            }
                            else
                            {
                                if (iElecObj.Core_Part_Number == "" && iElecObj.CableType != "X")
                                {
                                    modMain.arrFTcwob[rowNumber, 4] = $"{iElecObj.WireNumber}/{(!string.IsNullOrEmpty(iElecObj.Gauge) && iElecObj.Gauge.Length > 1 ? iElecObj.Gauge.Substring(1) : "")}";
                                }
                                else
                                {
                                    modMain.arrFTcwob[rowNumber, 4] = iElecObj.WireNumber + "/" + iElecObj.Gauge.Substring(1) + "/" + iElecObj.Core_Part_Number;
                                }
                            }
                            if (!string.IsNullOrEmpty(ConnectedEQU) && !string.IsNullOrEmpty(ConnectedPin))
                            {
                                modMain.arrFTcwob[rowNumber, 2] = ConnectedEQU;
                                modMain.arrFTcwob[rowNumber, 3] = ConnectedPin;
                            }
                            else
                            {
                                modMain.arrFTcwob[rowNumber, 2] = "No Connection";
                                modMain.arrFTcwob[rowNumber, 3] = "--NC--";
                            }
                            rowNumber++;
                        }
                        else
                        {
                            Logging.Error($"PrintConnectedLRU: rowNumber {rowNumber} exceeds arrFTcwob row limit.");
                            break;
                        }
                    }                  

                    modMain.SearchAndAppend(iElecObj.ConnectorName ?? "", ref modMain.arrListOfContinuityComponents);
                    if (!string.IsNullOrEmpty(ConnectedEQU))
                    {
                        modMain.SearchAndAppend(ConnectedEQU, ref modMain.arrListOfContinuityComponents);
                    }
                }
            }
            catch (Exception ex)
            {
                Logging.Error($"PrintConnectedLRU Exception: {ex.Message}\n{ex.StackTrace}");
            }
        }*/

        #endregion

        #region Old methods
        // old PrintConnectedLRU commented on July 2nd before implementing advanced error handling
        /* void PrintConnectedLRU(ElectreObject iElecObj, string istr1, ref int rowNumber)
         {
             // When multiple components are connected through a JM they all have to be listed in the collection
             // So created this function. It's common for both JM and BC

             string[] CollectiveValues;
             string[] TracedValue;
             string ConnectedEQU;
             string ConnectedPin;
             //int rowNumber = 0;


             // This section was added due to BC's connecting JMs and Splices
             CollectiveValues = istr1.Split(new[] { "%%" }, StringSplitOptions.None);

             for (int T = 1; T <= CollectiveValues.Length - 2; T++)
             {
                 TracedValue = CollectiveValues[T].Split('#');
                 ConnectedEQU = TracedValue[0];
                 ConnectedPin = TracedValue[1];

                 // Below condition is to avoid the same connection at From and To
                 // Content inside the condition is to put all the LRUs connected through JM to this iElecObj
                 string str9 = iElecObj.ConnectorName + iElecObj.PinNumber;
                 string str8 = ConnectedEQU + ConnectedPin;



                     if (str8 != str9)
                     {
                        modMain.arrFTcwob[rowNumber, modMain.FC1] = iElecObj.ConnectorName;
                        modMain.arrFTcwob[rowNumber, modMain.FP1] = iElecObj.PinNumber;
                        modMain.arrFTcwob[rowNumber, modMain.RDno] = iElecObj.SheetName;
                       modMain.arrFTcwob[rowNumber,        6    ] = modMain.layerAssign(iElecObj.Layer);

                     // Check if the CableType is in the specified string sCableType_O
                     if (modMain.sCableType_O.Contains("_" + iElecObj.CableType + "_"))
                      {
                            if (iElecObj.Core_Part_Number == "" && iElecObj.CableType != "X")
                            {
                                 modMain.arrFTcwob[rowNumber, modMain.WC] = $"{iElecObj.WireNumber}/{(!string.IsNullOrEmpty(iElecObj.Gauge) && iElecObj.Gauge.Length > 1 ? iElecObj.Gauge.Substring(1) : "")}";
                            }
                             else
                             {
                                 modMain.arrFTcwob[rowNumber, modMain.WC] = iElecObj.WireNumber + "/" + iElecObj.Gauge.Substring(1) + "/" + iElecObj.Core_Part_Number;
                             }

                      }
                      else if (modMain.sCableType_1.Contains($"_{iElecObj.CableType}_"))
                      {


                         modMain.arrFTcwob[rowNumber, modMain.WC] = $"{iElecObj.WireNumber}/{iElecObj.Core_Part_Number}";//{E1.Tag7} add in the between wirenumber and tag3 if required
                      }
                      else if (modMain.sCableType_2.Contains($"_{iElecObj.CableType}_"))
                      {
                         modMain.arrFTcwob[rowNumber, modMain.WC] = $"{iElecObj.WireNumber}/{iElecObj.Core_Part_Number}";//{E1.Tag7} add in the between wirenumber and tag3 if required
                      }
                      else
                      {

                             if (iElecObj.Core_Part_Number == "" && iElecObj.CableType != "X")
                             {
                                 modMain.arrFTcwob[rowNumber, modMain.WC] = $"{iElecObj.WireNumber}/{(!string.IsNullOrEmpty(iElecObj.Gauge) && iElecObj.Gauge.Length > 1 ? iElecObj.Gauge.Substring(1) : "")}";
                             }
                             else
                             {
                                 modMain.arrFTcwob[rowNumber, modMain.WC] = iElecObj.WireNumber + "/" + iElecObj.Gauge.Substring(1) + "/" + iElecObj.Core_Part_Number;
                             }

                      }




                      if (!string.IsNullOrEmpty(ConnectedEQU) && !string.IsNullOrEmpty(ConnectedPin))
                      {
                         modMain.arrFTcwob[rowNumber, modMain.TC1] = ConnectedEQU;
                         modMain.arrFTcwob[rowNumber, modMain.TP1] = ConnectedPin;
                      }
                      else
                      {
                         modMain.arrFTcwob[rowNumber, modMain.TC1] = "No Connection";
                         modMain.arrFTcwob[rowNumber, modMain.TP1] = "--NC--";
                      }

                       rowNumber++;
                     }

                 modMain.SearchAndAppend(iElecObj.ConnectorName, ref modMain.arrListOfContinuityComponents);
                 if (!string.IsNullOrEmpty(ConnectedEQU))
                 {
                     modMain.SearchAndAppend(ConnectedEQU, ref modMain.arrListOfContinuityComponents);
                 }

             }
         }*/

        // commented on June 3rd 2025 before implementing EQU double connector logic for swapping
        /* string TracePinOfBreakConnector(string iConnectorName, string iPinNumber, int iElecCollectionIndexSearched, string componentType)
         {
             // Variables declaration
             string BreakConnectorName; // If the given name is 104Dp (for pin side), this variable = 104D

             BreakConnectorName = iConnectorName.Substring(0, iConnectorName.Length - 1);

             // The array arrBCDetails is not used in this function, so it's omitted
             string GivenSide, MatchingSide; // M-Movable, F-Fixed
             ElectreObject b1;
             ElectreObject C1;
             string TracePinOfBreakConnectorResult = ""; // To accumulate the result

             // Get the last character of iConnectorName
             GivenSide = iConnectorName.Substring(iConnectorName.Length - 1);

             // Determine MatchingSide based on GivenSide
             if (GivenSide.ToLower() == "f")
             {
                 MatchingSide = "m";
             }
             else if (GivenSide.ToLower() == "m")
             {
                 MatchingSide = "f";
             }
             else
             {
                 MatchingSide = "";
             }

             // Other side of iConnectorName
             string ConnectingConnectorName = BreakConnectorName + MatchingSide;

             // Loop through ElecCollection
             for (int b = 0; b < selectedLoomObject.Count; b++)
             {
                 b1 = selectedLoomObject[b];
                 if (b1.ConnectorName.Equals(ConnectingConnectorName, StringComparison.OrdinalIgnoreCase) && b1.PinNumber == iPinNumber)
                 {
                     // Inner loop
                     for (int c = 0; c < selectedLoomObject.Count; c++)
                     {
                         C1 = selectedLoomObject[c];

                         // Check conditions
                         if (!string.IsNullOrEmpty(C1.WireNumber) &&
                             C1.WireNumber == b1.WireNumber &&
                             C1.SubNet == b1.SubNet &&
                             C1.ConnectorName != b1.ConnectorName)
                         {
                             switch (C1.ComponentType)
                             {
                                 case "EQU":
                                     TracePinOfBreakConnectorResult = C1.ConnectorName + "#" + C1.PinNumber + "%%" + TracePinOfBreakConnectorResult;
                                     c = selectedLoomObject.Count;

                                     if (modMain.connectorMap.Keys.Any(k => C1.ConnectorName.EndsWith(k, StringComparison.OrdinalIgnoreCase)))
                                     {
                                         TracePinOfBreakConnectorResult = TracePinOfBreakConnector(C1.ConnectorName, C1.PinNumber, c, C1.ComponentType) + TracePinOfBreakConnectorResult;
                                     }
                                     break;
                                 case "REL":
                                 case "SWT":
                                 case "ERM":
                                 case "IND":
                                 case "SCB":
                                 case "TCB":
                                 // case "TER":
                                 case "FUS":
                                 case "POT":
                                 case "LMP":
                                 case "ANT":
                                 case "BUS":
                                 case "MSW":
                                     TracePinOfBreakConnectorResult = C1.ConnectorName + "#" + C1.PinNumber + "%%" + TracePinOfBreakConnectorResult;
                                     // Exit the inner loop
                                     c = selectedLoomObject.Count;
                                     break;

                                 case "DIS":
                                     // Recursive call
                                     TracePinOfBreakConnectorResult = TracePinOfBreakConnector(C1.ConnectorName, C1.PinNumber, c, C1.ComponentType) + TracePinOfBreakConnectorResult;
                                     break;

                                 case "TBK":
                                 case "TER":
                                 case "SPL":
                                     // Call TracePinOfJM function (needs to be implemented)
                                     TracePinOfBreakConnectorResult = TracePinOfJM(C1) + TracePinOfBreakConnectorResult;
                                     break;
                             }
                         }
                     }
                 }
             }

             // If result is empty, assign default value
             if (string.IsNullOrEmpty(TracePinOfBreakConnectorResult))
             {

                 // This is to pass empty string, so this function doesn't crash
                 //TracePinOfBreakConnectorResult = "" + "#" + "" + "%%";
                 TracePinOfBreakConnectorResult = iConnectorName + "#" + iPinNumber + "%%";
             }
             string connectorName = TracePinOfBreakConnectorResult.Split('#')[0];
             // string pinNumber = TracePinOfBreakConnectorResult.Split('#')[1];
             string pinNumber = TracePinOfBreakConnectorResult.Split('#')[1].Split(new[] { "%%" }, StringSplitOptions.None)[0];
             if (!modMain.destinationConnectors.Contains($"{connectorName},{pinNumber}"))
             {
                 modMain.destinationConnectors.Add($"{connectorName},{pinNumber}");
             }
             Logging.Info("TraceOfPinBC: " + TracePinOfBreakConnectorResult);

             return TracePinOfBreakConnectorResult;
         }*/

        // commented on June 6 2025 due to lack of performance
        /* public string[,] removeRedundants(string[,] arrFTcwob)
         {
             string[,] arrFTcwobDist = new string[arrFTcwob.GetLength(0), 7];
             List<string> rededuntCheck = new List<string>();
             HashSet<string> duplicateDestTracker = new HashSet<string>();
             int arrFTcwobDistCount = 0;

             // Step 1: Count how many times each (DestConnector, DestPin) occurs
             Dictionary<string, int> destPinCounts = new Dictionary<string, int>();
             for (int i = 0; i < arrFTcwob.GetLength(0); i++)
             {
                 string destKey = $"{arrFTcwob[i, 2]}#{arrFTcwob[i, 3]}";
                 if (!destPinCounts.ContainsKey(destKey))
                     destPinCounts[destKey] = 0;
                 destPinCounts[destKey]++;
             }

             // Step 2: Loop and keep only first instance of each duplicate (2,3) and remove mirrored connections
             for (int i = 0; i < arrFTcwob.GetLength(0); i++)
             {
                 string destKey = $"{arrFTcwob[i, 2]}#{arrFTcwob[i, 3]}";

                 // If duplicate and already added once, skip
                 if (destPinCounts[destKey] > 1 && duplicateDestTracker.Contains(destKey))
                     continue;

                 string connectLine = $"{arrFTcwob[i, 2]}#{arrFTcwob[i, 3]}##{arrFTcwob[i, 0]}#{arrFTcwob[i, 1]}";
                 string presentLine = $"{arrFTcwob[i, 0]}#{arrFTcwob[i, 1]}##{arrFTcwob[i, 2]}#{arrFTcwob[i, 3]}";

                 // Check for reverse duplicates
                 if (!rededuntCheck.Contains(presentLine) && !rededuntCheck.Contains(connectLine))
                 {
                     for (int col = 0; col < 7; col++)
                         arrFTcwobDist[arrFTcwobDistCount, col] = arrFTcwob[i, col];

                     arrFTcwobDistCount++;
                     rededuntCheck.Add(presentLine);

                     // Mark this duplicate (2,3) pair as added once
                     if (destPinCounts[destKey] > 1)
                         duplicateDestTracker.Add(destKey);
                 }
             }

             return arrFTcwobDist;
         }*/

        /*private string[,] removeRedudents(string[,] arrFTcwob)
        {
            string[,] arrFTcwobDist;
            arrFTcwobDist = new string[arrFTcwob.GetLength(0), 7];
            List<string> rededuntCheck = new List<string>() ;
            int arrFTcwobDistCount = 0;

            for (int i = 0; i <arrFTcwob.GetLength(0); i++)
            {
                string connectLine = $"{arrFTcwob[i,2]}#{arrFTcwob[i, 3]}##{arrFTcwob[i, 0]}#{arrFTcwob[i, 1]}";
                string presentLine = $"{arrFTcwob[i, 0]}#{arrFTcwob[i, 1]}##{arrFTcwob[i, 2]}#{arrFTcwob[i, 3]}";

                if (!rededuntCheck.Contains(presentLine))//to skip the redunt line.
                {
                   
                    arrFTcwobDist[arrFTcwobDistCount, 0] = arrFTcwob[i, 0];
                    arrFTcwobDist[arrFTcwobDistCount, 1] = arrFTcwob[i, 1];
                    arrFTcwobDist[arrFTcwobDistCount, 2] = arrFTcwob[i, 2];
                    arrFTcwobDist[arrFTcwobDistCount, 3] = arrFTcwob[i, 3];
                    arrFTcwobDist[arrFTcwobDistCount, 4] = arrFTcwob[i, 4];
                    arrFTcwobDist[arrFTcwobDistCount, 5] = arrFTcwob[i, 5];
                    arrFTcwobDist[arrFTcwobDistCount, 6] = arrFTcwob[i, 6];
                    arrFTcwobDistCount++;
                    rededuntCheck.Add(connectLine);
                }
            }

            return arrFTcwobDist;
        }*/

        //for megger sheet3
        /*   public string[,] removeRedudent(string[,] arrFTcwob)
           {
               string[,] arrFTcwobDist;
               arrFTcwobDist = new string[arrFTcwob.GetLength(0), 4];
               List<string> rededuntCheck = new List<string>();
               int arrFTcwobDistCount = 0;

               for (int i = 0; i < arrFTcwob.GetLength(0); i++)
               {
                   string connectLine = $"{arrFTcwob[i, 2]}#{arrFTcwob[i, 3]}##{arrFTcwob[i, 0]}#{arrFTcwob[i, 1]}";
                   string presentLine = $"{arrFTcwob[i, 0]}#{arrFTcwob[i, 1]}##{arrFTcwob[i, 2]}#{arrFTcwob[i, 3]}";

                   if (!rededuntCheck.Contains(presentLine))//to skip the redunt line.
                   {

                       arrFTcwobDist[arrFTcwobDistCount, 0] = arrFTcwob[i, 0];
                       arrFTcwobDist[arrFTcwobDistCount, 1] = arrFTcwob[i, 1];
                       arrFTcwobDist[arrFTcwobDistCount, 2] = arrFTcwob[i, 2];
                       arrFTcwobDist[arrFTcwobDistCount, 3] = arrFTcwob[i, 3];

                       arrFTcwobDistCount++;
                       rededuntCheck.Add(connectLine);
                   }
               }

               return arrFTcwobDist;
           }*/

        /*public string[,] ConnectionComponentsforMegger()
        {
            ContinuityWOBreakDown();
            var listOfContinuityComponents = modMain.arrListOfContinuityComponents;
            if (listOfContinuityComponents == null || listOfContinuityComponents.Count == 0)
            {
                return new string[0, 0];  // Return an empty 2D array if the list is empty or null
            }

            // Define the number of rows and columns for the 2D array
            int rows = listOfContinuityComponents.Count;
            int cols = 1;  // Assuming you want a single column for each continuity component

            // Create a new 2D string array
            string[,] continuityComponentsArray = new string[rows, cols];

            // Fill the 2D array with data from the list
            for (int i = 0; i < rows; i++)
            {
                continuityComponentsArray[i, 0] = listOfContinuityComponents[i];  // Place each element in the first column
            }

            // Return the 2D array
            return continuityComponentsArray;
        }*/

        //for Megger sheet2 continuity
        /*  public string[,] ConnectionListforMegger()
          {

              selectedLoomObject = loomElctreObject(modMain.ElecCollection, _selectedLoomlist);

              selectedLoomObject = selectedLoomObject.OrderByDescending(obj => obj.ComponentType == "EQU") // "EQU" comes first
                                     .ThenBy(obj => obj.ConnectorName, StringComparer.OrdinalIgnoreCase) // Group by ConnectorName
                                     .ThenBy(obj => ExtractNumericPrefix(obj.PinNumber)) // Sort numerically if possible
                                     .ThenBy(obj => obj.PinNumber, StringComparer.OrdinalIgnoreCase) // Finally, sort lexicographically
                                     .ToList();

              modMain.ElecCollectionforNoMegger = selectedLoomObject.Where(e => e.NoMegger.Contains("No Megger")).ToList();

              selectedLoomObject = selectedLoomObject.Where(e => string.IsNullOrEmpty(e.NoMegger) || !e.NoMegger.Contains("No Megger")).ToList();

              //AssignInitialLink();
              ContinuityWOBreakDown();
              string[,] arrFTcwobDist = removeRedudent(modMain.arrFTcwob);
              return arrFTcwobDist;

          }*/

        /* public string ShuntOfConnector(string iToConnector, string iToPin, List<ElectreObject> electreCollection)
 {
     string shuntOfConnector = null;

     for (int j = 0; j < electreCollection.Count; j++)
     {
         var item = electreCollection[j];
         if (item.ConnectorName == iToConnector && item.PinNumber == iToPin &&
             !string.IsNullOrEmpty(item.WireNumber))
         {
             shuntOfConnector = item.Shunt;
             break;
         }
     }
     return shuntOfConnector;
 }*/

        /* public bool ValidateIfJumper(string iFromConnector, string iToConnector)
         {

             if (IfExist(iFromConnector, modMain.arrListOfJUNnSPL) && IfExist(iToConnector, modMain.arrListOfJUNnSPL))
             {
                 return true;
             }
             return false;
         }*/

        /*        public void AssignLinkToProjectWirelist()
        {
            for (int i = 0; i < selectedLoomObject.Count; i++)
            {
                ElectreObject electreObject = selectedLoomObject[i];
                if (electreObject.WireNumber != "" && electreObject.ComponentType.ToUpper() == "TBK") { }
            }
        }*/

        /*public void AssignInitialLink()
        {
            int a;
            string S1;
            string FC;
            string FP;
            string TC;
            string TP;
            string fs;
            string TS;
            string str1, str2;

            foreach (ElectreObject electreObject in selectedLoomObject)
            {
                if (!string.IsNullOrEmpty(electreObject.WireNumber)  || electreObject.ComponentType.ToUpper()=="SPL")
                {
                    S1 =electreObject.ConnectorName +";" + electreObject.ComponentType;

                    if (!modMain.SearchAndAppend(S1, ref modMain.arrAddress))
                    {
                        a = modMain.arrAddress.Count;
                        modMain.arrShuntExt1.Add("Link"+a);
                    }
                }
            }

            for (int kk = modMain.arrFT_CwithBCProject.GetLowerBound(0); kk <=modMain.arrFT_CwithBCProject.GetUpperBound(0); kk++)
            {

                if (_selectedLoomlist.Contains(modMain.arrFT_CwithBCProject[kk, 8]))
                {

                    FC = modMain.arrFT_CwithBCProject[kk, 0]?.ToString() ?? string.Empty;
                    FP = modMain.arrFT_CwithBCProject[kk, 1]?.ToString() ?? string.Empty;
                    fs = ShuntOfConnector(FC, FP, selectedLoomObject);
                    TC = modMain.arrFT_CwithBCProject[kk, 2]?.ToString() ?? string.Empty;
                    TP = modMain.arrFT_CwithBCProject[kk, 3]?.ToString() ?? string.Empty;

                    if (string.IsNullOrEmpty(FC) && string.IsNullOrEmpty(modMain.arrFT_CwithBCProject[kk, 2]?.ToString()))
                        break;

                    if (ValidateIfJumper(FC, TC))
                    {
                        TS = ShuntOfConnector(TC, TP, selectedLoomObject);
                        str1 = $"{FC};{fs}";
                        str2 = $"{TC};{TS}";

                        for (int aa = 0; aa <= modMain.arrAddress.Count; aa++)
                        {
                            if (modMain.arrAddress[aa] == str2)
                            {
                                modMain.arrShuntExt1[aa] = modMain.arrShuntExt1[ItemNumber(str1, modMain.arrAddress)];
                            }
                        }
                    }
                }

            }

        }*/

        /* public int ItemNumber(string istr, List<string> iarr)
         {
             for (int i = 0; i<iarr.Count; i++)
             {
                 if (iarr[i] == istr)
                 {
                     return i;
                 }
             }
             return -1;
         }*/

        #endregion
    }
}
