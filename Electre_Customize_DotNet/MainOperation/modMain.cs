using Electre_Customize_DotNet.Exceptions;
using Electre_Customize_DotNet.Forms;
using Electre_Customize_DotNet.Logs;
using Electre_Customize_DotNet.Objects;
using Electre_Customize_DotNet.Reports;
using Electre_Customize_DotNet.ReportUI;
using Microsoft.Office.Interop.Excel;
using System.Collections.Concurrent;
using System.Configuration;
using System.Reflection.Emit;
using System.Runtime.InteropServices.Marshalling;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Windows.Forms;
using static System.Runtime.InteropServices.JavaScript.JSType;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;
using Application = System.Windows.Forms.Application;
using Excel = Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;
using Electre_Customize_DotNet.Helpers.CableList;
using Electre_Customize_DotNet.Helpers.ExcelHelpers;
using Electre_Customize_DotNet.Helpers.Global;
using Electre_Customize_DotNet.Helpers.Megger;

namespace Electre_Customize_DotNet.MainOperation
{
    internal class modMain
    {
        // Global variables

        //organize the code its becoming too big, segregate into different classes
        public static List<ElectreObject> ElecCollection = new List<ElectreObject>();
        public static List<ElectreObject> ElecCollection_Panel = new List<ElectreObject>();
        public static List<ElectreObject> ElecCollection_All = new List<ElectreObject>();
        public static List<ElectreObject> ElecCollectionforNoMegger = new List<ElectreObject>();
        public static ElectreObject[] ElectreObjs;
        public static ElectreObject ElectreObj;
        public static string[,] arrFTcwob;       // Array of From and To - Without Break Connectors
        public static object[,] arrFT_CwithBC;   // Array of From and To - With Break Connectors
        public static object[,] arrFT_CwithBCProject;   // Array after removing duplicates
        public static object[,] arrFT_Cwithbundelproject;
        public static object[,] arrFT_CwithBCProjectPageOff;    // To handle page off separately
        public static object[,] arrFTcc;          // Array of From and To - Component connections
        public static int cwobReportRow; // Report row number
        public static int CwithBCReportRow;  // Report row number
        public static int CwithBCReportRowProject;   // Report row number
        public static int ccReportRow;   // Report row number for Component Connections
        public static int FC1; // From Connector1
        public static int FP1;   // From Pin1
        public static int TC1;   // To Connector1
        public static int TP1;   // To Pin1
        public static int WC;    // Wire code
        public static int RDno;  // Reference Drawing number
        public static int Shunt; // Shunt
        public static List<string> arrJunctionModule = new List<string>();
        public static List<string> arrEquipment = new List<string>();
        public static List<string> arrListOfComponents = new List<string>();
        public static List<string> arrListOfContinuityComponents = new List<string>();
        public static List<string> arrSpecificComponents = new List<string>();
        public static List<string> arrSubJMToConn = new List<string>();
        public static List<string> arrSubJMToPin = new List<string>();
        public static List<string> arrJMConnectionList = new List<string>();
        public static List<string> arrMatFCmod = new List<string>();
        public static List<string> arrMatFPmod = new List<string>();
        public static List<string> arrMatTCmod = new List<string>();
        public static List<string> arrMatTPmod = new List<string>();
        public static List<string> arrMatWireCode = new List<string>();
        public static List<string> arrMatRDnumber = new List<string>();
        public static List<string> arrListOfJUN = new List<string>();
        public static List<string> arrListOfEQU = new List<string>();
        public static List<string> arrListOfREL = new List<string>();
        public static List<string> arrListOfDIS = new List<string>();
        public static List<string> arrListOfSPL = new List<string>();
        public static List<string> arrListOfSWT = new List<string>();
        public static List<string> arrListOfLOOM = new List<string>();
        public static List<string> arrListOfSHEET = new List<string>();
        public static List<string> arrListOfPANEL = new List<string>();
        public static object[,] arrListOfComponentsInPANEL;
        public static object[,] arrMegger;
        public static object[,] arrMeggerYes;
        public static object[,] arrMeggerNo;
        public static object[,] arrMeggerYesLow;
        public static Library Library; //megger;
        public static List<Library> ElecList = new List<Library>(); //megger
        public static List<MeggerData> listmegger = new List<MeggerData>(); // formegger text file
        public static string[,] arrmegger;

        // For listing all pins in report
        public static List<string> arrConnectorNAME = new List<string>();
        public static List<string> arrConnectorPIN = new List<string>();
        public static List<string> arrConnectorNAMEandPIN = new List<string>();

        public static List<string> arrComponentsSelectedFromListsType = new List<string>();
        public static List<string> arrComponentsSelectedFromListsType2 = new List<string>();

        public static List<string> arrAddress = new List<string>();
        public static List<string> arrShuntExt1 = new List<string>();
        public static List<string> arrListOfJUNnSPL = new List<string>();
        public static object[,] arrSDSloom;
        public static int SDSloomCount;
        public static string sELECTRE_CUSTOMIZE;
        public static string sPROJECT_PATH;
        public static string sPROJECT_NAME;
        public static List<string> arrPowerOnCircuitBreakers = new List<string>();
        public static List<string> arrPowerOnCircuitBreakersType = new List<string>();
        public static List<string> arrFTpoweron = new List<string>();
        public static string[,] arrTableOfOOTBms_Panel;
        public static int OOTBmsTotalrow_Selection;
        public static string sCableType_O;   // String joining all the normal cables
        public static string sCableType_1;   // String joining the Special cables1
        public static string sCableType_2;   // String joining the Special cables2 ' CAT5 cables
        public static bool EXEinBatchMode;
        //public static int EXEReportOutputMode;
        public static bool CWOBactivate;

        private LoadingForm _loadingForm;
        private CustomReportWindow _customReportWindow;
        public static bool duplicatemessage = false;
        public static HashSet<string> partnumbersList = new HashSet<string>();
        // public static List<string> partnumbersandconnectorslist = new List<string>();
        public static Dictionary<string, HashSet<string>> partNumbersandConnectors = new Dictionary<string, HashSet<string>>();
        public static Dictionary<string, List<string>> connectorsAndPartNames = new Dictionary<string, List<string>>();
        public static List<string> destinationConnectors = new List<string>();
        public static string ChkListPanelText = string.Empty;
        public static ConcurrentBag<string> highMeggerData = new();
        public static List<string> lowMeggerData = new List<string>();
        public static List<(string ConnectorName, string PinNumber)> listConnectorPinsLibrary = new List<(string ConnectorName, string PinNumber)>();

        public modMain(CustomReportWindow customReportWindow, LoadingForm loadingForm)
        {
            _customReportWindow = customReportWindow;
            _loadingForm = loadingForm;
        }

        public static void MainFunction()
        {
            ElecCollection.Clear();
            try
            {
                // Variable declarations
               // bool bDel;
                ElectreObjs = new ElectreObject[1];
                int ObjCount = 0;
                string[] arrTemp;
                string[] arrCSVrow = new string[1];
                string[] iarrj = new string[1];
                int i;
                cwobReportRow = 0;
                CwithBCReportRow = 0;
                ccReportRow = 0;
                FC1 = 0;
                FP1 = 1;
                TC1 = 2;
                TP1 = 3;
                WC = 4;
                RDno = 5;
                Shunt = 6; // The same column number is used for showing the loom details also

                string EXEReportOutputMode = "0"; // "0" for GeneralExtraction, "1" for PanelExtraction

                using (StreamReader reader = new StreamReader(GlobalVar.DataExtractionGlobal))
                {
                    while (!reader.EndOfStream)
                    {
                        ObjCount++;
                        arrCSVrow[0] = reader.ReadLine();
                        arrTemp = arrCSVrow[0].Split(';');

                        if (!string.IsNullOrEmpty(arrTemp[12]))
                        {
                            if (EXEReportOutputMode == "1") // PanelExtraction
                            {
                                // Skip if no panel data in column 28
                                if (string.IsNullOrEmpty(arrTemp[28])) continue;
                            }
                            else if (EXEReportOutputMode == "0") // GeneralExtraction
                            {
                                if (!string.IsNullOrEmpty(arrTemp[28])) continue;
                            }

                            ElectreObject electreObj = new ElectreObject
                            {
                                SheetName = arrTemp[0],
                                SheetNumber = arrTemp[1],
                                DrawingNumber = arrTemp[2],
                                DefaultGauge = arrTemp[3],
                                BundleName = arrTemp[4],
                                EquipmentName = arrTemp[5],
                                ConnectorName = arrTemp[6],
                                PinNumber = arrTemp[7],
                                Ends = arrTemp[8],
                                FunctionalDesignation = arrTemp[9],
                                SymbolName = arrTemp[10],
                                ComponentType = arrTemp[11],
                                WireNumber = arrTemp[12],
                                Group = arrTemp[13],
                                Gauge = arrTemp[14],
                                CableType = arrTemp[15],
                                Length = arrTemp[16],
                                Signal = arrTemp[17],
                                Layer = arrTemp[18],
                                OverShield = arrTemp[19],
                                Net = arrTemp[21],
                                SubNet = arrTemp[22],
                                Shunt = arrTemp[23],
                                //Tag1 = arrTemp[25],
                                //Tag2 = "2", // To track linked connection state
                                Core_Part_Number = arrTemp[24], // Core/Part number For wire number in STP cases
                                //Tag4 = arrTemp[26], // Min Voltage
                                Voltage = arrTemp[27], // Max Voltage
                                Tag6_Link = "",
                                Tag7 = arrTemp[15], // Retain CableType for special cables
                                Panel = arrTemp[28],
                                ShuntExt1 = arrTemp[23],
                                NoMegger = arrTemp[29]
                                // Tag8 = arrTemp[30],
                            };

                            if (!string.IsNullOrEmpty(arrTemp[29]))
                            {
                                electreObj.CableType = arrTemp[29];
                            }

                            ElecCollection.Add(electreObj);
                        }

                        if (!string.IsNullOrEmpty(arrTemp[24]) && arrTemp[24].Length > 2)
                        {
                            if (!partnumbersList.Contains(arrTemp[24]))
                            {
                                partnumbersList.Add(arrTemp[24]);
                            }
                        }

                        if ((!string.IsNullOrEmpty(arrTemp[24]) && arrTemp[24].Length > 2) && !string.IsNullOrEmpty(arrTemp[06]))
                        {
                            if (!partNumbersandConnectors.ContainsKey(arrTemp[24]))
                            {
                                partNumbersandConnectors.Add(arrTemp[24], new HashSet<string> { arrTemp[06] });
                            }
                            else
                            {
                                partNumbersandConnectors[arrTemp[24]].Add(arrTemp[06]);
                            }
                        }

                        if ((!string.IsNullOrEmpty(arrTemp[24]) && arrTemp[24].Length > 2) && !string.IsNullOrEmpty(arrTemp[06]))
                        {
                            if (!connectorsAndPartNames.ContainsKey(arrTemp[06]))
                            {
                                //  connectorsAndPartNames.Add(arrTemp[06], arrTemp[24]);
                                connectorsAndPartNames[arrTemp[06]] = new List<string>();
                            }
                            connectorsAndPartNames[arrTemp[06]].Add(arrTemp[24]);
                        }
                    }
                }

                Logging.Info("Electre object List created");
                ElecCollection = ElecCollection.OrderBy(E => E.ConnectorName).ThenBy(e => e.PinNumber).ToList();                

                #region //Duplicate wire code commented on Dec 26, 2025
                /* List<WireListCount> wireListCountList1 = ElecCollection
                 .GroupBy(p => new { p.WireNumber })
                 .Where(g => g.Count() > 2)
                 .Select(g => new WireListCount()
                 {
                     WireNumber = g.Key.WireNumber,
                     // Core_Part_Number = g.Key.Core_Part_Number,
                     WireCount = g.Count(),
                 })
                 .ToList();

                 if (wireListCountList1.Any(e => e.WireCount > 2))
                 {
                     List<DuplicateWire> duplicateWires = new List<DuplicateWire>();

                     // Convert wireListCountList to a HashSet for faster lookup
                     var wireSet = new HashSet<string>(
                         wireListCountList1.Select(w => w.WireNumber.ToLower())
                     );

                     foreach (ElectreObject elec in ElecCollection)
                     {
                         if (!string.IsNullOrEmpty(elec.Core_Part_Number))
                         {
                             continue;
                         }

                         var key = elec.WireNumber.ToLower();
                         foreach (ElectreObject ele in ElecCollection)
                         {
                             if (wireSet.Contains(key) && !string.IsNullOrEmpty(ele.Core_Part_Number) && ele.WireNumber.ToLower() == key)
                             {
                                 duplicateWires.Add(new DuplicateWire
                                 {
                                     WireName = ele.WireNumber,
                                     WireNumber = ele.Core_Part_Number.ToLower(),
                                     sheetName = ele.SheetName
                                 });
                             }
                         }
                     }

                     var wireDuplicateFinale = duplicateWires
                         .GroupBy(g => new { g.WireName, g.sheetName, g.WireNumber })
                         .Select(g => new
                         {
                             g.Key.WireName,
                             g.Key.WireNumber,
                             g.Key.sheetName,
                             Count = g.Count() / 2,  // Assigning count properly
                         });

                     DuplWireLogging.DeletePreviousLogs();
                     foreach (var wire in wireDuplicateFinale)
                     {
                         DuplWireLogging.Info($"{wire.WireName}  {wire.sheetName} {wire.Count}");
                     }
                     duplicatemessage = true;

                     MessageBox.Show(
                         $"Duplicate Wires were found.\nPlease find the logs at {GlobalVar.StrtCmd}Logs\\DuplicateWire_{DateTime.Now:yyyy-MM-dd}.log",
                         "Duplicate Wires Found",
                         MessageBoxButtons.OK,
                         MessageBoxIcon.Exclamation
                     );

                     Application.Exit();
                 }*/
                #endregion

                // Array of cable types to differentiate regular and special cables
                string[] arrCableType = { "_", "PT", "PB", "PTB", "TB", "TTB", "QT", "QB", "QTB", "HTTP", "HTTQ", "HTTT", "HTSTP", "HTSTQ", "HTSTT", "SP",
                    "TP", "STP", "ST", "TT", "STT", "SQ", "TQ", "STQ", "HTS", "HTSS", "ST5","ST6", "ST7", "ST8","ST9", "ST10", "_" };//wirecode with - wireNumber & .Gauge & Subnet
                sCableType_O = string.Join("_", arrCableType);

                string[] arrCableType_1 = { "_", "X", "TX", "BX", "_" };  // Removed BC to add the core number in report
                sCableType_1 = string.Join("_", arrCableType_1);

                string[] arrCableType_2 = { "_", "CAT5STP", "CAT5STQ", "BC", "_" }; // Added BC to add the core number in report
                sCableType_2 = string.Join("_", arrCableType_2); //'Wirecode with - wireNumber & Subnet

                // Initialize arrays
                arrFTcwob = new string[ElecCollection.Count + 1, 7];  // used for Continuity
                arrFT_CwithBC = new object[ElecCollection.Count + 1, 11];  // used for Removing_DuplicateWires_In2DArray method
                                                                           // arrFT_CwithBCProject = new object[ElecCollection.Count + 1, 13];
                arrFT_CwithBCProject = new object[ElecCollection.Count + 1, 14];
                arrFT_CwithBCProjectPageOff = new object[ElecCollection.Count + 1, 10];
                arrFTcc = new object[ElecCollection.Count + 1, 8];
                arrMeggerYesLow = new object[301, 7];
                arrSDSloom = new object[501, 10];
                arrTableOfOOTBms_Panel = new string[3001, 81];
                SDSloomCount = 0;

                // Process ElectreObjects
                // arrListOfPANEL and arrConnectorNAMEandPIN are never sorted afterward in this
                // method (unlike the ten lists below), so their final order is whatever first-
                // occurrence insertion order UniqueStringList produces - they keep using it. The
                // other ten always get .OrderBy().ToList()'d right after this loop regardless of
                // insertion order, so accumulating them into a plain HashSet (dedup only, no
                // per-call scan/cache lookup at all) and sorting once at the end is equivalent and
                // cheaper than even the cached UniqueStringList.
                var setComponents = new HashSet<string>();
                var setLoom = new HashSet<string>();
                var setSheet = new HashSet<string>();
                var setJun = new HashSet<string>();
                var setJunSpl = new HashSet<string>();
                var setEqu = new HashSet<string>();
                var setRel = new HashSet<string>();
                var setDis = new HashSet<string>();
                var setSpl = new HashSet<string>();
                var setSwt = new HashSet<string>();

                for (i = 0; i < ElecCollection.Count; i++)
                {
                    ElectreObj = ElecCollection[i];

                    setComponents.Add(ElectreObj.ConnectorName);
                    setLoom.Add(ElectreObj.BundleName);
                    setSheet.Add(ElectreObj.SheetName);
                    UniqueStringList.TryAdd(ElectreObj.Panel, ref arrListOfPANEL);

                    switch (ElectreObj.ComponentType)
                    {
                        case "TBK":
                            setJun.Add(ElectreObj.ConnectorName);
                            setJunSpl.Add(ElectreObj.ConnectorName); // For JM Link
                            break;
                        case "EQU":
                            setEqu.Add(ElectreObj.ConnectorName); // For Connector
                            break;
                        case "REL":
                            setRel.Add(ElectreObj.ConnectorName);   // For Realy
                            break;
                        case "DIS":
                            setDis.Add(ElectreObj.ConnectorName);  // For Break Connector
                            break;
                        case "SPL":
                            setSpl.Add(ElectreObj.ConnectorName);
                            setJunSpl.Add(ElectreObj.ConnectorName); // For JM Link
                            break;
                        case "SWT":
                        case "ERM":
                        case "IND":
                        case "SCB":
                        case "TCB":
                        case "TER":
                        case "FUS":
                        case "POT":
                        case "LMP":
                        case "ANT":
                        case "BUS":
                        case "MSW":
                            setSwt.Add(ElectreObj.ConnectorName);
                            break;
                    }

                    // List the number of pins in each connector
                    if (ElectreObj.ComponentType == "EQU")
                    {
                        UniqueStringList.TryAdd($"{ElectreObj.ConnectorName};{ElectreObj.PinNumber}", ref arrConnectorNAMEandPIN);
                    }
                }

                arrListOfComponentsInPANEL = new object[setComponents.Count, 4];

                // sorting lists by order
                arrListOfLOOM = setLoom.OrderBy(x => x).ToList();
                arrListOfSHEET = setSheet.OrderBy(x => x).ToList();
                arrListOfComponents = setComponents.OrderBy(x => x).ToList();
                arrListOfJUN = setJun.OrderBy(x => x).ToList();
                arrListOfJUNnSPL = setJunSpl.OrderBy(x => x).ToList();
                arrListOfSPL = setSpl.OrderBy(x => x).ToList();
                arrListOfEQU = setEqu.OrderBy(x => x).ToList();
                arrListOfREL = setRel.OrderBy(x => x).ToList();
                arrListOfDIS = setDis.OrderBy(x => x).ToList();
                arrListOfSWT = setSwt.OrderBy(x => x).ToList();

                //Reading_And_StoringData_In2DArray();
                //Removing_DuplicateWires_In2DArray();
            }

            //catch (DuplicatWireException ex)
            //{
            //    throw new DuplicatWireException();
            //}

            catch (Exception ex)
            {
                Logging.Error("Issue Creating object list:" + ex.Message);
                throw new Exception();
            }
        }

        public static void MainFunction_withPanels()
        {
            // ElecCollection.Clear();
            try
            {
                arrListOfPANEL.Clear();
                // Variable declarations
                ElectreObjs = new ElectreObject[1];
                int ObjCount = 0;
                string[] arrTemp;
                string[] arrCSVrow = new string[1];
                string[] iarrj = new string[1];
                int i;
                cwobReportRow = 0;
                CwithBCReportRow = 0;
                ccReportRow = 0;
                FC1 = 0;
                FP1 = 1;
                TC1 = 2;
                TP1 = 3;
                WC = 4;
                RDno = 5;
                Shunt = 6; // The same column number is used for showing the loom details also

                // string EXEReportOutputMode = "0"; // "0" for GeneralExtraction, "1" for PanelExtraction

                using (StreamReader reader = new StreamReader(GlobalVar.DataExtractionGlobal))
                {
                    while (!reader.EndOfStream)
                    {
                        ObjCount++;
                        arrCSVrow[0] = reader.ReadLine();
                        arrTemp = arrCSVrow[0].Split(';');

                        if (!string.IsNullOrEmpty(arrTemp[12]))
                        {
                            if (string.IsNullOrEmpty(arrTemp[28])) continue;

                            ElectreObject electreObj = new ElectreObject
                            {
                                SheetName = arrTemp[0],
                                SheetNumber = arrTemp[1],
                                DrawingNumber = arrTemp[2],
                                DefaultGauge = arrTemp[3],
                                BundleName = arrTemp[4],
                                EquipmentName = arrTemp[5],
                                ConnectorName = arrTemp[6],
                                PinNumber = arrTemp[7],
                                Ends = arrTemp[8],
                                FunctionalDesignation = arrTemp[9],
                                SymbolName = arrTemp[10],
                                ComponentType = arrTemp[11],
                                WireNumber = arrTemp[12],
                                Group = arrTemp[13],
                                Gauge = arrTemp[14],
                                CableType = arrTemp[15],
                                Length = arrTemp[16],
                                Signal = arrTemp[17],
                                Layer = arrTemp[18],
                                OverShield = arrTemp[19],
                                Net = arrTemp[21],
                                SubNet = arrTemp[22],
                                Shunt = arrTemp[23],
                                // Tag1 = arrTemp[25],
                                //Tag2 = "2", // To track linked connection state
                                Core_Part_Number = arrTemp[24], // For wire number in STP cases
                                                                // Tag4 = arrTemp[26], // Min Voltage
                                Voltage = arrTemp[27], // Max Voltage
                                Tag6_Link = "",
                                Tag7 = arrTemp[15], // Retain CableType for special cables
                                Panel = arrTemp[28],
                                ShuntExt1 = arrTemp[23],
                                NoMegger = arrTemp[29]
                            };

                            if (!string.IsNullOrEmpty(arrTemp[29]))
                            {
                                electreObj.CableType = arrTemp[29];
                            }
                            ElecCollection_Panel.Add(electreObj);
                        }
                    }
                }
                Logging.Info("Electre object List created");
                ElecCollection_Panel = ElecCollection_Panel.OrderBy(E => E.ConnectorName).ThenBy(e => e.PinNumber).ToList();

                // Merge ElecCollection and ElecCollection_Panel into ElecCollection_All
                ElecCollection_All.AddRange(ElecCollection);
                ElecCollection_All.AddRange(ElecCollection_Panel);

                // Calling Duplicates wires checking method to check for duplicate wires in the entire project
                DuplicateWiresCheck.DuplicateWires();

                #region // Dupplicates wires code commented on Jan 16, 2026
                // Finding duplicate wires in the project
                //List<WireListCount> wireListCountList = ElecCollection_Panel
                //.GroupBy(p => new { p.WireNumber, p.Core_Part_Number })
                //.Where(g => g.Count() > 3)
                //.Select(g => new WireListCount()
                //{
                //    WireNumber = g.Key.WireNumber,
                //    Core_Part_Number = g.Key.Core_Part_Number,
                //    WireCount = g.Count(),
                //})
                //.ToList();

                //if (wireListCountList.Any(e => e.WireCount >= 3))
                //{
                //    List<DuplicateWire> duplicateWires = new List<DuplicateWire>();

                //    // Convert wireListCountList to a HashSet for faster lookup
                //    var wireSet = new HashSet<(string, string)>(
                //        wireListCountList.Select(w => (w.WireNumber.ToLower(), w.Core_Part_Number.ToLower()))
                //    );

                //    foreach (ElectreObject elec in ElecCollection_Panel)
                //    {
                //        var key = (elec.WireNumber.ToLower(), elec.Core_Part_Number.ToLower());

                //        if (wireSet.Contains(key))
                //        {
                //            duplicateWires.Add(new DuplicateWire
                //            {
                //                WireName = elec.WireNumber,
                //                WireNumber = elec.Core_Part_Number.ToLower(),
                //                sheetName = elec.SheetName
                //            });
                //        }
                //    }

                //    var wireDuplicateFinale = duplicateWires
                //        .GroupBy(g => new { g.WireName, g.sheetName, g.WireNumber })
                //        .Select(g => new
                //        {
                //            g.Key.WireName,
                //            g.Key.WireNumber,
                //            g.Key.sheetName,
                //            Count = g.Count() / 2,  // Assigning count properly
                //        });

                //    DuplWireLogging.DeletePreviousLogs();
                //    foreach (var wire in wireDuplicateFinale)
                //    {
                //        DuplWireLogging.Info($"{wire.WireName}  {wire.sheetName} {wire.Count}");
                //    }
                //    duplicatemessage = true;

                //    MessageBox.Show(
                //        $"Duplicate Wires were found in the project.\nPlease find the logs at {GlobalVar.StrtCmd}Logs\\DuplicateWire_{DateTime.Now:yyyy-MM-dd}.log",
                //        "Duplicate Wires Found",
                //        MessageBoxButtons.OK,
                //        MessageBoxIcon.Exclamation
                //    );

                //    Application.Exit();

                //    //throw new DuplicatWireException();
                //}
                #endregion

                // Array of cable types to differentiate regular and special cables
                string[] arrCableType = { "_", "PT", "PB", "PTB", "TB", "TTB", "QT", "QB", "QTB", "HTTP", "HTTQ", "HTTT", "HTSTP", "HTSTQ", "HTSTT", "SP",
             "TP", "STP", "ST", "TT", "STT", "SQ", "TQ", "STQ", "HTS", "HTSS", "ST5","ST6", "ST7", "ST8","ST9", "ST10", "_" };//wirecode with - wireNumber & .Gauge & Subnet
                sCableType_O = string.Join("_", arrCableType);

                string[] arrCableType_1 = { "_", "X", "TX", "BX", "_" };  // Removed BC to add the core number in report
                sCableType_1 = string.Join("_", arrCableType_1);

                string[] arrCableType_2 = { "_", "CAT5STP", "CAT5STQ", "BC", "_" }; // Added BC to add the core number in report
                sCableType_2 = string.Join("_", arrCableType_2); //'Wirecode with - wireNumber & .Tag7 & Subnet

                // Process ElectreObjects
                // arrListOfPANEL is sorted right after this loop (unlike in MainFunction(), where
                // it's left in insertion order), so it's safe to accumulate into a plain HashSet
                // here instead of the order-preserving UniqueStringList cache.
                int panelCount = ElecCollection_Panel.Count;
                var setPanel = new HashSet<string>(panelCount);
                var setBreakers = new HashSet<string>(arrPowerOnCircuitBreakers);
                var setBreakerTypes = new HashSet<string>(arrPowerOnCircuitBreakersType);

                for (i = 0; i < ElecCollection_Panel.Count; i++)
                {
                    ElectreObj = ElecCollection_Panel[i];
                    string connector = ElectreObj.ConnectorName;

                    if (ElectreObj.Panel != null)
                        setPanel.Add(ElectreObj.Panel);

                    switch (ElectreObj.ComponentType)
                    {
                        case "SCB":
                        case "TCB":
                            if (connector != null && setBreakers.Add(connector))
                                arrPowerOnCircuitBreakers.Add(connector);
                            if (ElectreObj.ComponentType != null && setBreakerTypes.Add(ElectreObj.ComponentType))
                                arrPowerOnCircuitBreakersType.Add(ElectreObj.ComponentType);
                            break;
                    }

                    // List the number of pins in each connector
                    if (ElectreObj.ComponentType == "EQU")
                    {
                        UniqueStringList.TryAdd($"{ElectreObj.ConnectorName};{ElectreObj.PinNumber}", ref arrConnectorNAMEandPIN);
                    }
                }               

                ElecCollection_All = ElecCollection_All.OrderBy(E => E.ConnectorName).ThenBy(e => e.PinNumber).ToList();

                arrListOfComponentsInPANEL = new object[arrListOfComponents.Count, 4];

                arrListOfPANEL = setPanel.OrderBy(x => x).ToList();

                Reading_And_StoringData_In2DArray();
                Removing_DuplicateWires_In2DArray();
            }

            //catch (DuplicatWireException ex)
            //{
            //    throw new DuplicatWireException();
            //}

            catch (Exception ex)
            {
                Logging.Error("Issue Creating object list:" + ex.Message);
                throw new Exception();
            }
        }

        //this method assigning data from Elecollection object to 2d array arrFT_CwithBC and this arrFT_CwithBC Used for Componentes specifice breakdown reports generation purpose
        public static void Removing_DuplicateWires_In2DArray()
        {
            try
            {
                CwithBCReportRow = 0;
                Logging.Info($"Removing_DuplicateWires_In2DArray started");

                // Same shape of bug as Reading_And_StoringData_In2DArray: the inner do-while
                // rescanned the ENTIRE collection for every i, making this O(ElecCollection.Count^2).
                // Unlike that method, this one has no "consume/Link" state - it deliberately emits
                // BOTH directions of every matching pair (i->j and, later, j->i), so the group of
                // candidates sharing a key must stay available for every member to enumerate, not
                // be drained like a queue. Group indices by (WireNumber, SubNet) once instead.
                var wireSubNetGroups = WireSubNetIndex.BuildGroups(ElecCollection);

                for (int i = 0; i < ElecCollection.Count; i++)
                {
                    ElectreObject E1 = ElecCollection[i];

                    if (!string.IsNullOrEmpty(E1.WireNumber) && E1.ComponentType != "SDS")
                    {
                        if (CwithBCReportRow > ElecCollection.Count)
                        {
                            break;
                            //MessageBox.Show("Same multiple wireCode found on the Project..\n Kindly resolve.", "Same WireCode", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            // MessageBox.Show("Elecollection don't have data");
                            //Application.Exit();
                        }
                        arrFT_CwithBC[CwithBCReportRow, FC1] = E1.ConnectorName;
                        arrFT_CwithBC[CwithBCReportRow, FP1] = E1.PinNumber;
                        arrFT_CwithBC[CwithBCReportRow, 5] = E1.Tag7;//Change from cableType 
                        arrFT_CwithBC[CwithBCReportRow, 6] = E1.Length;
                        arrFT_CwithBC[CwithBCReportRow, 7] = E1.SheetName;
                        arrFT_CwithBC[CwithBCReportRow, 8] = E1.BundleName;

                        string coreNumber = CoreNumberRules.IsValid(E1.Core_Part_Number) ? E1.Core_Part_Number : "";
                        // Build wire code

                        if (sCableType_O.Contains("_" + E1.CableType + "_") && E1.CableType != "HTSS")
                        {
                            if (coreNumber == "" && E1.CableType != "X")
                            {
                                arrFT_CwithBC[CwithBCReportRow, WC] = $"{E1.WireNumber}/{(!string.IsNullOrEmpty(E1.Gauge) && E1.Gauge.Length > 1 ? E1.Gauge.Substring(1) : "")}";
                            }
                            else
                            {
                                arrFT_CwithBC[CwithBCReportRow, WC] = $"{E1.WireNumber}/{(!string.IsNullOrEmpty(E1.Gauge) && E1.Gauge.Length > 1 ? E1.Gauge.Substring(1) + "/" : "")}/{coreNumber}";
                            }
                        }
                        else if (sCableType_1.Contains($"_{E1.CableType}_"))
                        {
                            arrFT_CwithBC[CwithBCReportRow, WC] = $"{E1.WireNumber}/{coreNumber}";//{E1.Tag7} add in the between wirenumber and tag3 if required
                        }
                        else if (sCableType_2.Contains($"_{E1.CableType}_"))
                        {
                            arrFT_CwithBC[CwithBCReportRow, WC] = $"{E1.WireNumber}/{coreNumber}";//{E1.Tag7} add in the between wirenumber and tag3 if required
                        }
                        else
                        {
                            if (coreNumber == "" && E1.CableType != "X")
                            {
                                arrFT_CwithBC[CwithBCReportRow, WC] = $"{E1.WireNumber}/{(!string.IsNullOrEmpty(E1.Gauge) && E1.Gauge.Length > 1 ? E1.Gauge.Substring(1) : "")}";
                            }
                            else
                            {
                                arrFT_CwithBC[CwithBCReportRow, WC] = $"{E1.WireNumber}/{(!string.IsNullOrEmpty(E1.Gauge) && E1.Gauge.Length > 1 ? E1.Gauge.Substring(1) + "/" : "")}{coreNumber}";
                            }

                        }

                        // Original scanned j=0..n-1 in order finding every match; replicate that
                        // ordering using the precomputed group instead of rescanning ElecCollection.
                        var key = WireSubNetIndex.KeyOf(E1);
                        if (wireSubNetGroups.TryGetValue(key, out var group))
                        {
                            foreach (int j in group)
                            {
                                if (j == i)
                                {
                                    continue; // matches the original's `j != i`
                                }

                                arrFT_CwithBC[CwithBCReportRow, TC1] = ElecCollection[j].ConnectorName;
                                arrFT_CwithBC[CwithBCReportRow, TP1] = ElecCollection[j].PinNumber;
                                arrFT_CwithBC[CwithBCReportRow, 10] = LayerCodes.ToNerd(ElecCollection[j].Layer);

                                CwithBCReportRow++;
                            }
                        }
                    }
                }
                Logging.Info($"Removing_DuplicateWires_In2DArray Ended");
            }

            catch (Exception ex)
            {
                Logging.Error(ex.Message);
                MessageBox.Show("modMain ##001: " + ex.Message);
            }
        }

        //this method assigning data from Elcollection object to 2d array arrFT_CwithBCProject and this arrFT_CwithBCProject used for sheets,project wirelist, looms reports generation purpose
        public static void Reading_And_StoringData_In2DArray()
        {
            try
            {
                ElectreObject E1;
                CwithBCReportRowProject = 0;

                Logging.Info("Reading_And_StoringData_In2DArray Started");

                // The inner search below looks for the first OTHER index sharing the same
                // (WireNumber, SubNet) key (case-insensitive) that hasn't been consumed yet -
                // originally a fresh linear scan of the whole collection for every outer i, making
                // this method O(ElecCollection.Count^2). Since every valid index eventually gets
                // marked Tag6_Link="Linked" (either as some earlier outer i's own row, or as an
                // earlier i's partner), by the time the outer loop reaches index i every valid
                // index below i is already linked - so the only indices the inner scan can ever
                // actually match are the ones still ahead of it. Precomputing a per-key queue of
                // candidate indices (in original order) lets each index be dequeued at most once
                // across the whole run, turning the search amortized O(1) per outer i instead of
                // O(ElecCollection.Count).
                var partnerQueueByKey = WireSubNetIndex.BuildQueues(ElecCollection);

                for (int i = 0; i < ElecCollection.Count; i++)
                {
                    E1 = ElecCollection[i];
                    if (!string.IsNullOrEmpty(E1.WireNumber) && E1.ComponentType != "SDS" && E1.Tag6_Link != "Linked")
                    {
                        if (CwithBCReportRowProject > ElecCollection.Count)
                        {
                            break;
                            // MessageBox.Show(" Same multiple wireCode found on the Project..\n Kindly resolve.", "Same WireCode", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            //MessageBox.Show("Elecollection don't have data");
                            //Application.Exit();
                        }

                        E1.Tag6_Link = "Linked";
                        arrFT_CwithBCProject[CwithBCReportRowProject, FC1] = E1.ConnectorName;
                        arrFT_CwithBCProject[CwithBCReportRowProject, FP1] = E1.PinNumber;
                        arrFT_CwithBCProject[CwithBCReportRowProject, 5] = E1.Tag7;//Change from cableType 
                        arrFT_CwithBCProject[CwithBCReportRowProject, 6] = E1.Length;
                        arrFT_CwithBCProject[CwithBCReportRowProject, 7] = E1.SheetName;
                        arrFT_CwithBCProject[CwithBCReportRowProject, 8] = E1.BundleName;

                        string coreNumber = CoreNumberRules.IsValid(E1.Core_Part_Number) ? E1.Core_Part_Number : "";

                        if (sCableType_O.Contains($"_{E1.CableType}_") && E1.CableType != "HTSS")
                        {
                            if (coreNumber == "" && E1.CableType != "X")
                            {
                                arrFT_CwithBCProject[CwithBCReportRowProject, WC] = $"{E1.WireNumber}/{(!string.IsNullOrEmpty(E1.Gauge) && E1.Gauge.Length > 1 ? E1.Gauge.Substring(1) : "")}";
                            }
                            else
                            {
                                arrFT_CwithBCProject[CwithBCReportRowProject, WC] = $"{E1.WireNumber}/{(!string.IsNullOrEmpty(E1.Gauge) && E1.Gauge.Length > 1 ? E1.Gauge.Substring(1) + "/" : "")}{coreNumber}";
                            }

                        }
                        else if (sCableType_1.Contains($"_{E1.CableType}_"))
                        {
                            arrFT_CwithBCProject[CwithBCReportRowProject, WC] = $"{E1.WireNumber}/{coreNumber}";//{E1.Tag7} add in the between wirenumber and tag3 if required
                        }
                        else if (sCableType_2.Contains($"_{E1.CableType}_"))
                        {
                            arrFT_CwithBCProject[CwithBCReportRowProject, WC] = $"{E1.WireNumber}/{coreNumber}";//{E1.Tag7} add in the between wirenumber and tag3 if required
                        }
                        else
                        {
                            if (coreNumber == "" && E1.CableType != "X")
                            {
                                arrFT_CwithBCProject[CwithBCReportRowProject, WC] = $"{E1.WireNumber}/{(!string.IsNullOrEmpty(E1.Gauge) && E1.Gauge.Length > 1 ? E1.Gauge.Substring(1) : "")}";
                            }
                            else
                            {
                                arrFT_CwithBCProject[CwithBCReportRowProject, WC] = $"{E1.WireNumber}/{(!string.IsNullOrEmpty(E1.Gauge) && E1.Gauge.Length > 1 ? E1.Gauge.Substring(1) + "/" : "")}{coreNumber}";
                            }

                        }

                        // NOTE: the original has no `break` here - if more than one other index
                        // shares this key, EVERY one of them (not just the first) gets treated as
                        // a match and consumes its own row, same as reproduced below.
                        var key = WireSubNetIndex.KeyOf(E1);
                        if (partnerQueueByKey.TryGetValue(key, out var candidateQueue))
                        {
                            while (candidateQueue.Count > 0)
                            {
                                int j = candidateQueue.Dequeue();
                                if (j == i || ElecCollection[j].Tag6_Link == "Linked")
                                {
                                    continue; // matches the original's `i != j` / Tag6_Link check
                                }

                                arrFT_CwithBCProject[CwithBCReportRowProject, TC1] = ElecCollection[j].ConnectorName;
                                arrFT_CwithBCProject[CwithBCReportRowProject, TP1] = ElecCollection[j].PinNumber;
                                arrFT_CwithBCProject[CwithBCReportRowProject, 9] = ElecCollection[j].SheetName;
                                arrFT_CwithBCProject[CwithBCReportRowProject, 10] = ElecCollection[j].BundleName;
                                ElecCollection[j].Tag6_Link = "Linked";
                                arrFT_CwithBCProject[CwithBCReportRowProject, 12] = LayerCodes.ToNerd(ElecCollection[j].Layer);
                                arrFT_CwithBCProject[CwithBCReportRowProject, 13] = ElecCollection[j].Group;

                                CwithBCReportRowProject++;
                            }
                        }

                    }
                }

                Logging.Info("Reading_And_StoringData_In2DArray completed");
            }

            catch (Exception ex)
            {
                MessageBox.Show("modMain #002: " + ex.Message);
            }
        }
        //this method is the entry point for reports generations 
        public async void ReportExcecution()
        {
            string templpath = Path.Combine(GlobalVar.StrtCmd, "templ");

            string wireListLoomFolder = ConfigurationManager.AppSettings["LoomFolder"];
            string wireListSheetFolder = ConfigurationManager.AppSettings["SheetFolder"];
            string wireListBundleFolder = ConfigurationManager.AppSettings["BundleFolder"];
            string EquipmentFolder = ConfigurationManager.AppSettings["EquipmentFolder"];
            string BrkFolder = ConfigurationManager.AppSettings["BrkFolder"];
            string JmFolder = ConfigurationManager.AppSettings["JmFolder"];
            string MiscFolder = ConfigurationManager.AppSettings["MiscFolder"];
            string ContinuityFolder = ConfigurationManager.AppSettings["ContinuityFolder"];
            string ContinuityCompFolder = ConfigurationManager.AppSettings["ContinuityCompFolder"];
            string MeggerFolder = ConfigurationManager.AppSettings["MeggerSchedulerFolder"];
            string PowerOnFolder = ConfigurationManager.AppSettings["PowerOnFolder"];

            string PanelLoomFolderCL = ConfigurationManager.AppSettings["LoomFolderCL"];
            string PanelLoomFolderBundleCL = ConfigurationManager.AppSettings["LoomFolderBundleCL"];
            string PanelSheetFolderCL = ConfigurationManager.AppSettings["SheetFolderCL"];
            string PanelContinuityFolder = ConfigurationManager.AppSettings["panelContinuityFolder"];
            string PanelContinuityCompFolder = ConfigurationManager.AppSettings["panelContinuityCompFolder"];
            string PanelMeggerFolder = ConfigurationManager.AppSettings["panelMeggerFolder"];
            string PanelPowerOnFolder = ConfigurationManager.AppSettings["panelPowerOnFolder"];
            

            string errorMsg = "One or more paths are not configured for reports, please check the Configuration file";
            if (string.IsNullOrEmpty(wireListLoomFolder) || string.IsNullOrEmpty(wireListSheetFolder) || string.IsNullOrEmpty(wireListBundleFolder) ||
                string.IsNullOrEmpty(EquipmentFolder) || string.IsNullOrEmpty(BrkFolder) || string.IsNullOrEmpty(JmFolder) ||string.IsNullOrEmpty(MiscFolder) ||
                string.IsNullOrEmpty(ContinuityFolder) || string.IsNullOrEmpty(ContinuityCompFolder) || string.IsNullOrEmpty(MeggerFolder) ||
                string.IsNullOrEmpty(PowerOnFolder) || string.IsNullOrEmpty(PowerOnFolder) || string.IsNullOrEmpty(PanelLoomFolderCL) || 
                string.IsNullOrEmpty(PanelLoomFolderBundleCL) || string.IsNullOrEmpty(PanelSheetFolderCL) || string.IsNullOrEmpty(PanelContinuityFolder) || 
                string.IsNullOrEmpty(PanelContinuityCompFolder) || string.IsNullOrEmpty(PanelMeggerFolder) || string.IsNullOrEmpty(PanelPowerOnFolder) )
            {
                Logging.Error(errorMsg);
                MessageBox.Show(errorMsg, "Configuration Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
                //return;
            }

            if (_customReportWindow.SelectedPanelListPanelDrawing.Count > 0 && string.IsNullOrEmpty(PanelDrawingWindow.panelDigit))
            {
                MessageBox.Show($"Panel digit is not specified. Please Specify to Continue", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                _loadingForm.Close();
                return;
            }

            // Check first if any report is selected — exit early if none
            if (!_customReportWindow.projectWireList &&
                _customReportWindow.SelectedLoomList.Count == 0 &&
                _customReportWindow.SelectedSheetList.Count == 0 &&
                _customReportWindow.SelectedEqupList.Count == 0 &&
                _customReportWindow.SelectedBrkList.Count == 0 &&
                _customReportWindow.SelectedJmList.Count == 0 &&
                _customReportWindow.SelectedMiscList.Count == 0 &&
                _customReportWindow.SelectedContLoomList.Count == 0 &&
                _customReportWindow.SelectedMeggerLoomList.Count == 0 &&
                !_customReportWindow.powerOnProjectList &&
                _customReportWindow.SelectedPanelList.Count == 0 &&
                _customReportWindow.SelectedPanelListPanelDrawing.Count ==0 &&
                // newly added for Panel Drawing Schedules
                _customReportWindow.SelectedLoomListPanelDwgCL.Count == 0 &&
                _customReportWindow.SelectedSheetListPanelDwgCL.Count == 0 &&
                _customReportWindow.SelectedLoomListPanelDwgCont.Count == 0 &&
                _customReportWindow.SelectedLoomListPanelDwgMeg.Count == 0 &&
                _customReportWindow.SelectedSheetListPanelDwgPower.Count == 0)
            {
                _loadingForm.Close();
                MessageBox.Show(
                    "Please select at least one report type to generate.",
                    "No Report Selected",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                );
                return; // early exit, don't proceed
            }

            else
            {
                // _loadingForm.Show();
                //  Application.DoEvents();
                modExcel modExcelInst = new modExcel();
                modExcelInst.InitiateExcel();

                object[,] arrCustomLoom;

                // Disable the button and update the initial status on the loading form
                _loadingForm.Loadingbtn = "Wait";
                _loadingForm.LoadingbtnEnable = false;
                _loadingForm.LbsLoadinMessag = "Initializing Report Generation...";
                Application.DoEvents(); // Force UI updates immediately

                await Task.Run(() =>
                {
                    // Cable List
                    if (_customReportWindow.projectWireList)
                    {
                        var __swWireList = PerfDiagnostics.StartTimer();
                        CwithBCReportRow = 0;
                        Logging.Info("WireList report Started");
                        string workbookPath = modExcelInst.CreateNewWorkbook("WireList", "WireList", ConfigurationManager.AppSettings["WireListFolder"]);

                        if (!modExcelInst.CreateWirelistReportHeader(workbookPath))
                        {
                            Logging.Error("Failed to create the wireList Report Header");
                        }
                        WireListReport(modExcelInst);

                        // Update loading message on the UI thread
                        _loadingForm.Invoke(() =>
                        {
                            _loadingForm.LbsLoadinMessag = "WireList Report Generated";
                        });

                        Logging.Info("WireList report Completed");
                        PerfDiagnostics.StopAndLog(__swWireList, "ReportGeneration", "WireList", ("InputCount", ElecCollection_All?.Count ?? 0));
                    }

                    if (_customReportWindow.SelectedLoomList.Count > 0)
                    {
                        var __swLoom = PerfDiagnostics.StartTimer();
                        string fullPath = Path.Combine(templpath, wireListLoomFolder);
                        string fullPathBundle = Path.Combine(templpath, wireListBundleFolder);

                        //DeleteExistingFiles(fullPath);
                        //DeleteExistingFiles(fullPathBundle);

                        Incorrect_cable_group_IDs.DeletePreviousLogs();
                        Null_Group_IDs.DeletePreviousLogs();
                        loomReportWireList(modExcelInst, _customReportWindow.SelectedLoomList, wireListLoomFolder);
                        loomReportWireListBundle(modExcelInst, _customReportWindow.SelectedLoomList, wireListBundleFolder);

                        // Update loading message on the UI thread
                        _loadingForm.Invoke(() =>
                        {
                            _loadingForm.LbsLoadinMessag = "Looms Report Generated";
                        });
                        PerfDiagnostics.StopAndLog(__swLoom, "ReportGeneration", "LoomWireList", ("SelectedLoomCount", _customReportWindow.SelectedLoomList.Count));
                    }

                    if (_customReportWindow.SelectedSheetList.Count > 0)
                    {
                        var __swSheet = PerfDiagnostics.StartTimer();
                        string fullPath = Path.Combine(templpath, wireListSheetFolder);
                       // DeleteExistingFiles(fullPath);

                        sheetsReportWireList(modExcelInst, _customReportWindow.SelectedSheetList, wireListSheetFolder);

                        // Update loading message on the UI thread
                        _loadingForm.Invoke(() =>
                        {
                            _loadingForm.LbsLoadinMessag = "Sheet Report Generated";
                        });
                        PerfDiagnostics.StopAndLog(__swSheet, "ReportGeneration", "SheetWireList", ("SelectedSheetCount", _customReportWindow.SelectedSheetList.Count));
                    }

                    // Component Breakdown
                    if (_customReportWindow.SelectedEqupList.Count > 0)
                    {
                        var __swEqup = PerfDiagnostics.StartTimer();
                        string fullPath = Path.Combine(templpath, EquipmentFolder);
                        //DeleteExistingFiles(fullPath);

                        equipmentReportWireList(modExcelInst, "COMP_WireList", ConfigurationManager.AppSettings["EquipmentFolder"], _customReportWindow.SelectedEqupList);

                        // Update loading message on the UI thread
                        _loadingForm.Invoke(() =>
                        {
                            _loadingForm.LbsLoadinMessag = "Equipment Report Generated";
                        });
                        PerfDiagnostics.StopAndLog(__swEqup, "ReportGeneration", "Equipment", ("SelectedCount", _customReportWindow.SelectedEqupList.Count));
                    }

                    if (_customReportWindow.SelectedBrkList.Count > 0)
                    {
                        var __swBrk = PerfDiagnostics.StartTimer();
                        string fullPath = Path.Combine(templpath, BrkFolder);
                        //DeleteExistingFiles(fullPath);

                        equipmentReportWireList(modExcelInst, "BRK_WireList", ConfigurationManager.AppSettings["BrkFolder"], _customReportWindow.SelectedBrkList);

                        // Update loading message on the UI thread
                        _loadingForm.Invoke(() =>
                        {
                            _loadingForm.LbsLoadinMessag = "Break Down Report Generated";
                        });
                        PerfDiagnostics.StopAndLog(__swBrk, "ReportGeneration", "BreakConn", ("SelectedCount", _customReportWindow.SelectedBrkList.Count));
                    }

                    if (_customReportWindow.SelectedJmList.Count > 0)
                    {
                        var __swJm = PerfDiagnostics.StartTimer();
                        string fullPath = Path.Combine(templpath, JmFolder);
                        //DeleteExistingFiles(fullPath);

                        equipmentReportWireList(modExcelInst, "JM_WireList", ConfigurationManager.AppSettings["JmFolder"], _customReportWindow.SelectedJmList);

                        // Update loading message on the UI thread
                        _loadingForm.Invoke(() =>
                        {
                            _loadingForm.LbsLoadinMessag = "Junction Module Report Generated";
                        });
                        PerfDiagnostics.StopAndLog(__swJm, "ReportGeneration", "JunctionModule", ("SelectedCount", _customReportWindow.SelectedJmList.Count));
                    }

                    if (_customReportWindow.SelectedMiscList.Count > 0)
                    {
                        var __swMisc = PerfDiagnostics.StartTimer();
                        string fullPath = Path.Combine(templpath, MiscFolder);
                        //DeleteExistingFiles(fullPath);

                        equipmentReportWireList(modExcelInst, "MISC_WireList", ConfigurationManager.AppSettings["MiscFolder"], _customReportWindow.SelectedMiscList);

                        // Update loading message on the UI thread
                        _loadingForm.Invoke(() =>
                        {
                            _loadingForm.LbsLoadinMessag = "MISC Report Generated";
                        });
                        PerfDiagnostics.StopAndLog(__swMisc, "ReportGeneration", "Misc", ("SelectedCount", _customReportWindow.SelectedMiscList.Count));
                    }

                    //Continuity
                    if (_customReportWindow.SelectedContLoomList.Count > 0)
                    {
                        var __swCont = PerfDiagnostics.StartTimer();
                        ContinuityWOBreakdown cwob = new ContinuityWOBreakdown(modExcelInst, _customReportWindow.SelectedContLoomList);
                        cwob.ContinuityReportGeneration(ContinuityFolder, ContinuityCompFolder, ElecCollection);
                        // Update loading message on the UI thread
                        _loadingForm.Invoke(() =>
                        {
                            _loadingForm.LbsLoadinMessag = "Continuity Report Generated";
                        });
                        PerfDiagnostics.StopAndLog(__swCont, "ReportGeneration", "Continuity", ("SelectedLoomCount", _customReportWindow.SelectedContLoomList.Count), ("InputCount", ElecCollection?.Count ?? 0));
                    }

                    // Megger
                    if (_customReportWindow.SelectedMeggerLoomList.Count > 0)
                    {
                        var __swMegger = PerfDiagnostics.StartTimer();
                        MeggerSchedulesheet(modExcelInst, _customReportWindow.SelectedMeggerLoomList, MeggerFolder, ElecCollection);

                        // Update loading message on the UI thread
                        _loadingForm.Invoke(() =>
                        {
                            _loadingForm.LbsLoadinMessag = "Megger Report Generated";
                        });
                        PerfDiagnostics.StopAndLog(__swMegger, "ReportGeneration", "Megger", ("SelectedLoomCount", _customReportWindow.SelectedMeggerLoomList.Count));
                    }

                    // Power On
                    if (_customReportWindow.powerOnProjectList)
                    {
                        var __swPowerOn = PerfDiagnostics.StartTimer();
                        CwithBCReportRow = 0;
                        Logging.Info("Power On report Started");
                        string workbookPath = modExcelInst.CreateNewWorkbook("PowerOn_Project", "PowerOn", ConfigurationManager.AppSettings["PowerOnFolder"]);

                        if (!modExcelInst.CreatePowerOnReportHeader(workbookPath))
                        {
                            Logging.Error("Failed to create the PowerOn Report Header");
                        }

                        modExcelInst.AppendToExcelPowerOn(ElecCollection_All);
                        // modExcelInst.SortWirelistWireNumber(2, arrFT_CwithBCProject.GetLength(0) + 1);

                        // Update loading message on the UI thread
                        _loadingForm.Invoke(() =>
                        {
                            _loadingForm.LbsLoadinMessag = "PowerOn_Project Report Generated";
                        });

                        Logging.Info("PowerOn_Project report Completed");
                        PerfDiagnostics.StopAndLog(__swPowerOn, "ReportGeneration", "PowerOn_Project", ("InputCount", ElecCollection_All?.Count ?? 0));
                    }

                    if (_customReportWindow.SelectedPanelList.Count > 0)
                    {
                        var __swPowerOnPanel = PerfDiagnostics.StartTimer();
                        string fullPath = Path.Combine(templpath, PowerOnFolder);
                        //DeleteExistingFiles(fullPath);

                        var SelectedPanels = _customReportWindow.SelectedPanelList;
                        foreach (var panel in SelectedPanels)
                        {
                            string sanitizedSheetName = SheetNames.Sanitize(panel);
                            string workbookPath = modExcelInst.CreateNewWorkbook($"PowerOn_Panel_{panel}", $"Panel_{sanitizedSheetName}", PowerOnFolder);

                            if (!modExcelInst.CreatePowerOnReportHeader(workbookPath))
                            {
                                Logging.Error($"Failed to create the PowerOn Report Header for Panel {panel}");
                            }
                            PowerOnReport powerOn = new PowerOnReport(modExcelInst, new List<string> { panel });
                            powerOn.PowerOnReportGeneration();

                          /*  _loadingForm.Invoke(() =>
                            {
                                _loadingForm.LbsLoadinMessag = $"Panel {panel} Report Generated";
                            });*/
                        }
                        _loadingForm.Invoke(() =>
                        {
                            _loadingForm.LbsLoadinMessag = $"Power On Reports Generated";
                        });
                        PerfDiagnostics.StopAndLog(__swPowerOnPanel, "ReportGeneration", "PowerOn_Panel", ("SelectedPanelCount", SelectedPanels.Count));
                    }

                    // panel drawing
                    if (_customReportWindow.SelectedPanelListPanelDrawing.Count > 0)
                    {
                        ChkListPanelText = _customReportWindow._panelDrawingWindow.ChkListPanelText;

                        string panelDrawingFolderPath = ConfigurationManager.AppSettings["PanelDrawingFolder"];
                        string templFolder = Path.Combine(GlobalVar.ReportFolderGlobal, panelDrawingFolderPath);

                        if (Directory.Exists(templFolder))
                        {
                            string[] txtFiles = Directory.GetFiles(templFolder, "*.txt");

                            foreach (string file in txtFiles)
                            {
                                try
                                {
                                    File.Delete(file);
                                }
                                catch (IOException ex)
                                {
                                    MessageBox.Show($"Unable to delete file '{file}': {ex.Message}");
                                }
                                catch (UnauthorizedAccessException ex)
                                {
                                    MessageBox.Show($"Permission denied for file '{file}': {ex.Message}");
                                }
                            }
                        }

                        if (!Directory.Exists(templFolder))
                        {
                            Directory.CreateDirectory(templFolder);
                        }

                        // Create the PanelDetails file
                        string panelDetailsFilePath = Path.Combine(templFolder, PanelDrawingWindow.panelDigit + "-PanelDetails.txt");
                        File.Create(panelDetailsFilePath).Close();

                        // Create the ComponentsList file
                        string ComponentsListFilePath = Path.Combine(templFolder, PanelDrawingWindow.panelDigit + "-PanelComponentsList.txt");
                        File.Create(ComponentsListFilePath).Close();
          
                        List<string> materialsList = ElecCollection_All.Where(e => e.Panel == ChkListPanelText)
                                                                        .Select(e => e.ConnectorName)
                                                                        .Distinct().ToList();

                        //create and Append to PanelInfo txt
                        modExcelInst.AppendPanelInfoTxt(templFolder, ChkListPanelText);

                        // Append to material list txt
                        modExcelInst.AppendMaterialListTxt(templFolder, ChkListPanelText, ComponentsListFilePath, materialsList);

                        modExcelInst.AppendPanelDetailsTxt(templFolder, ChkListPanelText, panelDetailsFilePath);
                       

                        _loadingForm.Invoke(() =>
                        {
                            _loadingForm.LbsLoadinMessag = "Panel Reports Generated";
                        });
                    }

                    #region Panel drawing Schedules

                    // Panel Schedules Cable List
                    if (_customReportWindow.SelectedLoomListPanelDwgCL.Count > 0)
                    {
                        var __swPanelCL = PerfDiagnostics.StartTimer();
                        string fullPath = Path.Combine(templpath, PanelLoomFolderCL);
                        string fullPathBundle = Path.Combine(templpath, PanelLoomFolderBundleCL);

                        //DeleteExistingFiles(fullPath);
                        //DeleteExistingFiles(fullPathBundle);

                        Incorrect_cable_group_IDs.DeletePreviousLogs();
                        Null_Group_IDs.DeletePreviousLogs();
                        loomReportWireList(modExcelInst, _customReportWindow.SelectedLoomListPanelDwgCL, PanelLoomFolderCL, arrListOfSHEET);
                        loomReportWireListBundle(modExcelInst, _customReportWindow.SelectedLoomListPanelDwgCL, PanelLoomFolderBundleCL, arrListOfSHEET);

                        // Update loading message on the UI thread
                        _loadingForm.Invoke(() =>
                        {
                            _loadingForm.LbsLoadinMessag = "Looms Report Generated";
                        });
                        PerfDiagnostics.StopAndLog(__swPanelCL, "ReportGeneration", "PanelSchedules_CableList", ("SelectedLoomCount", _customReportWindow.SelectedLoomListPanelDwgCL.Count));
                    }

                    if (_customReportWindow.SelectedSheetListPanelDwgCL.Count > 0)
                    {
                        var __swPanelSheet = PerfDiagnostics.StartTimer();
                        string fullPath = Path.Combine(templpath, PanelSheetFolderCL);
                        //DeleteExistingFiles(fullPath);

                        sheetsReportWireList(modExcelInst, _customReportWindow.SelectedSheetListPanelDwgCL, PanelSheetFolderCL);

                        // Update loading message on the UI thread
                        _loadingForm.Invoke(() =>
                        {
                            _loadingForm.LbsLoadinMessag = "Sheet Report Generated";
                        });
                        PerfDiagnostics.StopAndLog(__swPanelSheet, "ReportGeneration", "PanelSchedules_Sheet", ("SelectedSheetCount", _customReportWindow.SelectedSheetListPanelDwgCL.Count));
                    }

                    // Panel Schedules Continuity
                    if (_customReportWindow.SelectedLoomListPanelDwgCont.Count > 0)
                    {
                        var __swPanelCont = PerfDiagnostics.StartTimer();
                        ContinuityWOBreakdown cwob = new ContinuityWOBreakdown(modExcelInst, _customReportWindow.SelectedLoomListPanelDwgCont);
                        cwob.ContinuityReportGeneration(PanelContinuityFolder, PanelContinuityCompFolder, PanelDrawingSchedulesWindow.filteredSheetCollection);
                        // Update loading message on the UI thread
                        _loadingForm.Invoke(() =>
                        {
                            _loadingForm.LbsLoadinMessag = "Continuity Report Generated";
                        });
                        PerfDiagnostics.StopAndLog(__swPanelCont, "ReportGeneration", "PanelSchedules_Continuity", ("SelectedLoomCount", _customReportWindow.SelectedLoomListPanelDwgCont.Count));
                    }

                    // Panel Schedules Megger
                    if (_customReportWindow.SelectedLoomListPanelDwgMeg.Count > 0)
                    {
                        var __swPanelMeg = PerfDiagnostics.StartTimer();
                        MeggerSchedulesheet(modExcelInst, _customReportWindow.SelectedLoomListPanelDwgMeg, PanelMeggerFolder, PanelDrawingSchedulesWindow.filteredSheetCollection);

                        // Update loading message on the UI thread
                        _loadingForm.Invoke(() =>
                        {
                            _loadingForm.LbsLoadinMessag = "Megger Report Generated";
                        });
                        PerfDiagnostics.StopAndLog(__swPanelMeg, "ReportGeneration", "PanelSchedules_Megger", ("SelectedLoomCount", _customReportWindow.SelectedLoomListPanelDwgMeg.Count));
                    }

                    // Panel Schedules Power On
                    if (_customReportWindow.SelectedSheetListPanelDwgPower.Count > 0)
                    {
                        var __swPanelPower = PerfDiagnostics.StartTimer();
                        string fullPath = Path.Combine(templpath, PanelPowerOnFolder);
                        if (!Directory.Exists(fullPath))
                        {
                            Directory.CreateDirectory(fullPath);
                        }
                       // DeleteExistingFiles(fullPath);

                        var SelectedSheets = _customReportWindow.SelectedSheetListPanelDwgPower;
                        foreach (var sheet in SelectedSheets)
                        {
                            string sanitizedSheetName = SheetNames.Sanitize(sheet);
                            string workbookPath = modExcelInst.CreateNewWorkbook($"PowerOn_Sheet_{sheet}", $"{sanitizedSheetName}", PanelPowerOnFolder);

                            if (!modExcelInst.CreatePowerOnReportHeader(workbookPath))
                            {
                                Logging.Error($"Failed to create the PowerOn Report Header for Sheet {sheet}");
                            }
                            PowerOnReport powerOn = new PowerOnReport(modExcelInst, new List<string> { sheet });
                            powerOn.PowerOnReportGeneration_PanelDWG();
                        }

                        _loadingForm.Invoke(() =>
                        {
                            _loadingForm.LbsLoadinMessag = $"Power On Reports Generated";
                        });
                        PerfDiagnostics.StopAndLog(__swPanelPower, "ReportGeneration", "PanelSchedules_PowerOn", ("SelectedSheetCount", SelectedSheets.Count));
                    }
                    #endregion

                });

                // Update the loading form UI after all operations are complete
                _loadingForm.Loadingbtn = "Close";
                _loadingForm.LoadingbtnEnable = true;
                _loadingForm.CompletionMessage = "Report Generation Completed";
                _loadingForm.ReportLoclbl = "Reports are Generated @" + GlobalVar.ReportFolderGlobal;

                // Release Excel resources
                modExcelInst.releaseExcel();
            }
        }

        // This method generates Excel reports for sheets in Cable List.
        private void sheetsReportWireList(modExcel modExcelInst, List<string> selectedSheetList, string sheetFolderPath)
        {
            try
            {
                //ElectreObject E17;
                // The buffer used to be allocated at the full height of the project array for every sheet. It now
                // only grows to this sheet's rows (+ blank rows, see RowBuffer), capped at that same height.
                int maxBufferRows = arrFT_CwithBC.GetUpperBound(0);
                Logging.Info("sheet wise Report generation Started");

                ColumnRowIndex rowsBySheet = null;   // built on first use, from column 9 (sheet name)
                object rowsBySheetLock = new object();

                // Every sheet is its own workbook. With App.config ExcelWorkerCount > 1 several Excel instances work through the
                // sheets at the same time; the default (1) is the old sequential loop on the shared instance. Each worker owns its
                // row buffer (the same reuse / reset rules as before, per worker).
                ExcelParallel.ForEach(selectedSheetList, modExcelInst, () => new WirelistBufferState(10), (excel, s, buffer) =>
                {
                    try
                    {
                        Logging.Info($"{s} Sheet report Started");
                       // bool appendPage = false;

                        ColumnRowIndex index;
                        lock (rowsBySheetLock)
                        {
                            rowsBySheet ??= new ColumnRowIndex(arrFT_CwithBCProject, 9);
                            index = rowsBySheet;
                        }
                        var sheetRows = index.RowsFor(s);
                        RowBuffer.EnsureCapacity(ref buffer.Rows, buffer.Z + sheetRows.Count, maxBufferRows);

                        foreach (int y in sheetRows)
                        {
                            // Copy elements from arrFT_CwithBCProject to arrCustomLoom
                            for (int x = 0; x < 9; x++)
                            {
                                buffer.Rows[buffer.Z, x] = arrFT_CwithBCProject[y, x];
                            }
                            buffer.Rows[buffer.Z, 9] = arrFT_CwithBCProject[y, 12];
                            buffer.Z++;
                        }

                        string sheetName = SheetNames.Sanitize(s);
                        // string workbookPath = modExcelInst.CreateNewWorkbook($"{s}_SHEET_Wirelist", sheetName, ConfigurationManager.AppSettings["SheetFolder"]);
                        if (ExcelSwitches.OnePassWirelistWorkbook)
                        {
                            excel.WriteWirelistWorkbookOnePass($"{s}_SHEET_Wirelist", sheetName, sheetFolderPath, buffer.Rows, maxBufferRows);
                        }
                        else
                        {
                            string workbookPath = excel.CreateNewWorkbook($"{s}_SHEET_Wirelist", sheetName, sheetFolderPath);

                            if (!excel.CreateWirelistReportHeader(workbookPath))
                            {
                                Logging.Error($"Error Sheet report Header {s}");
                            }

                            excel.AppendToExcel(buffer.Rows, saveWorkbook: false);
                            excel.SortWirelistWireNumber(2, maxBufferRows, saveBeforeClose: false);
                        }
                        // same request as before (the full height); SortWirelistWireNumber trims it to the written rows
                        buffer.Z = 0;
                        buffer.Rows = new object[0, 10];
                        Logging.Info($"{s} Sheet report Completed");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"{s}:{ex.Message}");
                        Logging.Error($"{s}:{ex.Message}");
                    }
                });
            }

            catch (Exception ex)
            {
                MessageBox.Show("modMain #003: " + ex.Message);
            }
        }

        //this method gennerates Excel reports for project WireList
        private void WireListReport(modExcel modExcelInst)
        {
            // The buffer used to be arrFT_CwithBCProject.Length rows (rows x columns!) of which only the populated prefix was ever written
            // (AppendToExcel stops at the first block of three blank rows). RowBuffer.EquivalentHeight gives the smallest height that
            // makes AppendToExcel write exactly the same rows.
            int sourceRows = arrFT_CwithBCProject.GetLength(0);
            int bufferRows = RowBuffer.EquivalentHeight(arrFT_CwithBCProject, arrFT_CwithBCProject.Length);
            object[,] arrCustomLoom = new object[bufferRows, 10];
            int Z = 0;

            for (int Y = 0; Y < Math.Min(sourceRows, bufferRows); Y++)
            {
                // Check if the value at position (Y, 9) matches the selected item in frmMain.listLOOM at index S
                for (int X = 0; X < 9; X++)  // Copy elements from arrFT_CwithBCProject to arrCustomLoom
                {
                    arrCustomLoom[Z, X] = arrFT_CwithBCProject[Y, X];
                }

                arrCustomLoom[Z, 9] = arrFT_CwithBCProject[Y, 12];

                Z++;
            }
            modExcelInst.AppendToExcel(arrCustomLoom, saveWorkbook: false);
            modExcelInst.SortWirelistWireNumber(2, arrFT_CwithBCProject.GetLength(0) + 1, saveBeforeClose: false);

        }

        //this method generates Excel reports for Component Specific Breakdown 
        private void equipmentReportWireList(modExcel modExcelInst, string fileSuffixName, string FolderName, List<string> componentList)
        {
            try
            {
                ElectreObject E17;

                // First component: the original buffer was arrFT_CwithBCProject.Length rows; later components got
                // arrFT_CwithBC.GetUpperBound(0) rows. Same growth rule as the sheet reports, with the cap of each case.
                int maxBufferRowsFirst = arrFT_CwithBCProject.Length;
                int maxBufferRowsNext = arrFT_CwithBC.GetUpperBound(0);

                ColumnRowIndex rowsByComponent = null;   // built on first use, from column 0 (component / connector)
                object rowsByComponentLock = new object();

                // Every component is its own workbook. With App.config ExcelWorkerCount > 1 several Excel instances work through the
                // components at the same time; the default (1) is the old sequential loop on the shared instance. Each worker owns
                // its row buffer and its height cap (first item: ...First, afterwards: ...Next), like the old locals.
                //foreach (string s in _mainForm.ListEQU2)
                ExcelParallel.ForEach(componentList, modExcelInst, () => new WirelistBufferState(10, maxBufferRowsFirst), (excel, s, buffer) =>
                {
                    try
                    {
                        Logging.Info($"{s} Equipment report Started");

                        ColumnRowIndex index;
                        lock (rowsByComponentLock)
                        {
                            rowsByComponent ??= new ColumnRowIndex(arrFT_CwithBC, 0);
                            index = rowsByComponent;
                        }
                        var componentRows = index.RowsFor(s);
                        RowBuffer.EnsureCapacity(ref buffer.Rows, buffer.Z + componentRows.Count, buffer.MaxRows);

                        foreach (int y in componentRows)
                        {
                            // Copy elements from arrFT_CwithBC to arrCustomLoom
                            for (int x = 0; x < 9; x++)
                            {
                                buffer.Rows[buffer.Z, x] = arrFT_CwithBC[y, x];
                            }
                            buffer.Rows[buffer.Z, 9] = arrFT_CwithBC[y, 10];
                            buffer.Z++;
                        }

                        if (ExcelSwitches.OnePassWirelistWorkbook)
                        {
                            excel.WriteWirelistWorkbookOnePass($"{s}_{fileSuffixName}", s, FolderName, buffer.Rows, buffer.MaxRows);
                        }
                        else
                        {
                            string workbookPath = excel.CreateNewWorkbook($"{s}_{fileSuffixName}", s, FolderName);

                            if (!excel.CreateWirelistReportHeader(workbookPath))
                            {
                                Logging.Error($"Error Report Header {s}");
                            }

                            excel.AppendToExcel(buffer.Rows, saveWorkbook: false);
                            excel.SortWirelistWireNumber(2, buffer.MaxRows, saveBeforeClose: false);
                        }
                        // same request as before (the full height); SortWirelistWireNumber trims it to the written rows
                        buffer.Z = 0;
                        buffer.Rows = new object[0, 10];
                        buffer.MaxRows = maxBufferRowsNext;
                        Logging.Info($"{s} Equipment report Completed");
                    }
                    catch (Exception ex)
                    {

                        MessageBox.Show($"{s}:{ex.Message}");
                        Logging.Error($"{s}:{ex.Message}");

                    }

                });
            }

            catch (Exception ex)
            {
                MessageBox.Show("modMain ##05: " + ex.Message);
            }

        }

        //this method generates Excel reports for without HAL Template Loom
        private void loomReportWireList(modExcel modExcelInst, List<string> selectedLoomList, string loomFolderPath, List<string> selectedSheets = null)
        {
            try
            {
                var loomRows = LoomRowIndex.Build(arrFT_CwithBCProject, selectedSheets);

                // Every loom is its own workbook: with App.config ExcelWorkerCount > 1 several Excel instances work through the looms at
                // the same time; the default (1) is the old sequential loop on the shared instance.
                //foreach (string loomName in _mainForm.ListLOOM)
                ExcelParallel.ForEach(selectedLoomList, modExcelInst, (excel, loomName) =>
                {
                    try
                    {
                        Logging.Info($"{loomName} Loom report without template Started");
                        // Sized to exactly this loom's rows. It used to be arrFT_CwithBCProject.Length rows (rows x columns)
                        // x 14 columns for EVERY loom - hundreds of MB of empty array that the loops below then scanned.
                        // Only the filled rows were ever used, and the loops skip empty rows, so the result is the same.
                        loomRows.TryGetRows(loomName, out var matchingRows);
                        object[,] arrCustomLoom = new object[matchingRows?.Count ?? 0, 14];
                        int Z = 0;

                        if (matchingRows != null)
                        {
                            foreach (var (Y, isFrom) in matchingRows)
                            {
                                if (isFrom)  // matches the original's arrFT_CwithBCProject[Y, 8] == loomName branch
                                {
                                    for (int X = 0; X < 9; X++)  // Copy elements from arrFT_CwithBCProject to arrCustomLoom
                                    {
                                        arrCustomLoom[Z, X] = arrFT_CwithBCProject[Y, X];
                                    }

                                    arrCustomLoom[Z, 9] = arrFT_CwithBCProject[Y, 12];
                                    arrCustomLoom[Z, 10] = arrFT_CwithBCProject[Y, 13];

                                    Z++;
                                }
                                else  // matches the original's arrFT_CwithBCProject[Y, 10] == loomName branch
                                {
                                    for (int X = 0; X < 9; X++)  // Copy elements from arrFT_CwithBCProject to arrCustomLoom
                                    {
                                        arrCustomLoom[Z, X] = arrFT_CwithBCProject[Y, X];
                                    }

                                    arrCustomLoom[Z, 9] = arrFT_CwithBCProject[Y, 12];
                                    arrCustomLoom[Z, 10] = arrFT_CwithBCProject[Y, 13];
                                    // Swap values as in the VBA code
                                    string tempFromConnector = arrCustomLoom[Z, 0]?.ToString();
                                    string tempFromPin = arrCustomLoom[Z, 1]?.ToString();

                                    arrCustomLoom[Z, 0] = arrCustomLoom[Z, 2];  // Swap position 1 with 3
                                    arrCustomLoom[Z, 1] = arrCustomLoom[Z, 3];  // Swap position 2 with 4
                                    arrCustomLoom[Z, 2] = tempFromConnector;
                                    arrCustomLoom[Z, 3] = tempFromPin;

                                    // Set specific columns with values from arrFT_CwithBCProject
                                    arrCustomLoom[Z, 7] = arrFT_CwithBCProject[Y, 9];
                                    arrCustomLoom[Z, 8] = arrFT_CwithBCProject[Y, 10];

                                    Z++;
                                }
                            }
                        }

                        List<WireListLoomSort> loomSortList = new List<WireListLoomSort>();
                        loomSortList.Clear();
                        WireListLoomSort loomObject;
                        for (int i = 0; i < arrCustomLoom.GetLength(0); i++)
                        {
                            if (!string.IsNullOrEmpty((string)arrCustomLoom[i, 4]))
                            {
                                loomObject = new WireListLoomSort();
                                loomObject.FromConn = (string)arrCustomLoom[i, 0];
                                loomObject.FromPin = (string)arrCustomLoom[i, 1];
                                loomObject.ToConn = (string)arrCustomLoom[i, 2];
                                loomObject.ToPin = (string)arrCustomLoom[i, 3];
                                loomObject.WireCode = (string)arrCustomLoom[i, 4];
                                loomObject.WireType = (string)arrCustomLoom[i, 5];
                                loomObject.length = (string)arrCustomLoom[i, 6];
                                loomObject.Sheet = (string)arrCustomLoom[i, 7];
                                loomObject.Bundle = (string)arrCustomLoom[i, 8];
                                loomObject.layer = (string)arrCustomLoom[i, 9];
                                loomObject.Group = (string)arrCustomLoom[i, 10];
                                loomSortList.Add(loomObject);
                            }
                        }

                        // LINQ query to exclude groups that not start with an alphabet
                        var combinedList = LoomGroupRules.FindIncorrectGroups(loomSortList, x => x.Group, x => x.WireCode);

                        if (combinedList.Count > 0)
                        {
                            foreach (var item in combinedList)
                            {
                                Incorrect_cable_group_IDs.Info($" LoomName: {loomName},Group Code :{item.Group}, Wire Code : {item.WireCode}");
                            }
                            MessageBox.Show($"{loomName}: Incorrect Cable Group Id");
                        }
                        //if group number null
                        var nullGroupList = loomSortList.Where(x => x.Group == null || x.Group == "").ToList();

                        foreach (var nullgroup in nullGroupList)
                        {
                            Null_Group_IDs.Info($" Wire Code: {nullgroup.WireCode}, Loom Name: {nullgroup.Bundle} , Sheet Name: {nullgroup.Sheet}");
                        }

                        // Drop groups starting with a letter, then order by group number / letters / sub-number / wire code
                        loomSortList = LoomGroupRules.FilterAndSort(loomSortList, x => x.Group, x => x.WireCode);

                        object[,] arrFinalReportLOOM = new object[loomSortList.Count, 11];
                        int m = 0;
                        for (int i = 0; i < loomSortList.Count; i++)
                        {
                            arrFinalReportLOOM[m, 0] = loomSortList[m].FromConn;
                            arrFinalReportLOOM[m, 1] = loomSortList[m].FromPin;
                            arrFinalReportLOOM[m, 2] = loomSortList[m].ToConn;
                            arrFinalReportLOOM[m, 3] = loomSortList[m].ToPin;
                            arrFinalReportLOOM[m, 4] = loomSortList[m].WireCode;
                            arrFinalReportLOOM[m, 5] = loomSortList[m].WireType;
                            arrFinalReportLOOM[m, 6] = loomSortList[m].length;
                            arrFinalReportLOOM[m, 7] = loomSortList[m].Sheet;
                            arrFinalReportLOOM[m, 8] = loomSortList[m].Bundle;
                            arrFinalReportLOOM[m, 9] = loomSortList[m].layer;
                            arrFinalReportLOOM[m, 10] = loomSortList[m].Group;
                            m++;
                        }

                        // string workbookPath = modExcelInst.CreateNewWorkbook($"{loomName}_Loom_Wirelist", loomName, ConfigurationManager.AppSettings["LoomFolder"]);
                        string workbookPath = excel.CreateNewWorkbook($"{loomName}_Loom_Wirelist", loomName, loomFolderPath);

                        if (!excel.CreateloomlistReportHeader(workbookPath))
                        {
                            Logging.Error($"Error Sheet report Header {loomName}");
                        }

                        excel.AppendToExcel(arrFinalReportLOOM);
                        excel.CloseReportWorkbookWithoutSaving();   // saved by AppendToExcel; it used to stay open until Excel quit (one per loom)
                        // modExcelInst.SortWirelistWireNumber3(2, arrFinalReportLOOM.GetUpperBound(0) + 1);
                        Logging.Info($"{loomName} Loom report without Template Completed");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"{loomName}:incorrect cable group id");
                        Logging.Error($"{loomName}: {ex.Message}");
                    }
                });
            }
            catch (Exception ex)
            {

                Logging.Error(ex.Message);
                MessageBox.Show("ModMain ##6" + ex.Message);
            }
        }

        // this method generates Excel reports for with HAL Template Loom
        private void loomReportWireListBundle(modExcel modExcelInst, List<string> selectedLoomList, string bundleFolderPath, List<string> selectedSheets = null)
        {
            string gug = string.Empty;

            try
            {
                // Only ever used via .Contains() below (inside a loom-rows loop, itself inside the
                // per-loom loop), so a HashSet avoids an O(allspecialcables.Count) scan per row.
                var allspecialcables = new HashSet<string>(ElecCollection.Where(p => p.CableType == "X").Select(p => p.Tag7));

                List<string> gaugeList = ElecCollection
                .Select(p => p.Gauge)
                    .Where(g => !string.IsNullOrEmpty(g)) // Ensure no null or empty values
                  .Distinct()
                 .ToList();

                // Check if gaugeList has elements before accessing index 0
                if (gaugeList.Any())
                {
                    string gauge = gaugeList[0]; // No need for ToString() since it's already a string
                    gug = gauge.Length > 0 ? gauge[0].ToString() : string.Empty; // Avoid IndexOutOfRange
                }


                var loomRows = LoomRowIndex.Build(arrFT_CwithBCProject, selectedSheets);

                // Every loom is its own workbook: with App.config ExcelWorkerCount > 1 several Excel instances work through the looms at
                // the same time; the default (1) is the old sequential loop on the shared instance.
                //foreach (string loomName in _mainForm.ListLOOM)
                ExcelParallel.ForEach(selectedLoomList, modExcelInst, (excel, loomName) =>
                {
                    try
                    {
                        Logging.Info($"{loomName} Loom report Started");
                        // Sized to exactly this loom's rows. It used to be arrFT_CwithBCProject.Length rows (rows x columns)
                        // x 14 columns for EVERY loom - hundreds of MB of empty array that the loops below then scanned.
                        // Only the filled rows were ever used, and the loops skip empty rows, so the result is the same.
                        loomRows.TryGetRows(loomName, out var matchingRows);
                        object[,] arrCustomLoom = new object[matchingRows?.Count ?? 0, 14];
                        int Z = 0;

                        if (matchingRows != null)
                        {
                            foreach (var (Y, isFrom) in matchingRows)
                            {
                                if (isFrom)  // matches the original's arrFT_CwithBCProject[Y, 8] == loomName branch
                                {
                                    for (int X = 0; X < 13; X++)  // Copy elements from arrFT_CwithBCProject to arrCustomLoom
                                    {
                                        arrCustomLoom[Z, X] = arrFT_CwithBCProject[Y, X];
                                    }
                                    arrCustomLoom[Z, 13] = arrFT_CwithBCProject[Y, 13];

                                    Z++;
                                }
                                else  // matches the original's arrFT_CwithBCProject[Y, 10] == loomName branch
                                {
                                    for (int X = 0; X < 9; X++)  // Copy elements from arrFT_CwithBCProject to arrCustomLoom
                                    {
                                        arrCustomLoom[Z, X] = arrFT_CwithBCProject[Y, X];
                                    }
                                    arrCustomLoom[Z, 9] = arrFT_CwithBCProject[Y, 12];
                                    arrCustomLoom[Z, 12] = arrFT_CwithBCProject[Y, 12];
                                    arrCustomLoom[Z, 13] = arrFT_CwithBCProject[Y, 13];

                                    Z++;
                                }
                            }
                        }

                        List<loomSort> loomSortList = new List<loomSort>();
                        loomSortList.Clear();
                        loomSort loomObject;

                        for (int i = 0; i < arrCustomLoom.GetLength(0); i++)
                        {
                            if (!string.IsNullOrEmpty((string)arrCustomLoom[i, 4]))
                            {
                                loomObject = new loomSort();
                                loomObject.FromConn = (string)arrCustomLoom[i, 0];
                                loomObject.FromPin = (string)arrCustomLoom[i, 1];
                                loomObject.ToConn = (string)arrCustomLoom[i, 2];
                                loomObject.ToPin = (string)arrCustomLoom[i, 3];
                                loomObject.WireCode = (string)arrCustomLoom[i, 4];
                                loomObject.WireType = (string)arrCustomLoom[i, 5];
                                loomObject.length = (string)arrCustomLoom[i, 6];
                                loomObject.Sheet = (string)arrCustomLoom[i, 7];
                                loomObject.Bundle = (string)arrCustomLoom[i, 8];

                                loomObject.layer = (string)arrCustomLoom[i, 12];
                                loomObject.Group = (string)arrCustomLoom[i, 13];

                                string[] wireCodeDisect = loomObject.WireCode.Split('/');

                                if (wireCodeDisect.Length == 3)
                                {
                                    if (string.IsNullOrEmpty(wireCodeDisect[2]))
                                    {
                                        loomObject.wireNumber = 1;
                                        loomObject.WireName = wireCodeDisect[0] + wireCodeDisect[1];
                                    }
                                    else
                                    {
                                        int wn = 1;
                                        loomObject.WireName = wireCodeDisect[0] + wireCodeDisect[1];
                                        if (int.TryParse(wireCodeDisect[2], out wn))
                                        {
                                            loomObject.wireNumber = wn;
                                        }
                                        else
                                        {
                                            loomObject.wireNumber = wn;
                                        }
                                    }
                                }
                                else if (wireCodeDisect.Length == 2)
                                {
                                    if (string.IsNullOrEmpty(wireCodeDisect[1]))
                                    {
                                        loomObject.wireNumber = 1;
                                        loomObject.WireName = wireCodeDisect[0];
                                    }
                                    else
                                    {
                                        int wn = 1;
                                        loomObject.WireName = wireCodeDisect[0];
                                        loomObject.wireNumber = wn;
                                    }
                                }
                                else if (wireCodeDisect.Length == 1)
                                {
                                    loomObject.wireNumber = 1;
                                    loomObject.WireName = wireCodeDisect[0];
                                }
                                else
                                {
                                    loomObject.wireNumber = 0;
                                    loomObject.WireName = "";
                                }
                                loomSortList.Add(loomObject);
                            }
                        }


                        // Drop groups starting with a letter, then order by group number / letters / sub-number / wire code
                        loomSortList = LoomGroupRules.FilterAndSort(loomSortList, x => x.Group, x => x.WireCode);

                        int loomRowCount = 0;

                        loomRowCount = loomSortList
                        .Where(x => !string.IsNullOrEmpty(x.FromConn) && !string.IsNullOrEmpty(x.WireCode))
                           .Count();

                        object[,] arrFinalReportLOOM = new object[loomRowCount + 1, 20];//incremented by one to assoiciate for the excel update.

                        loomSort arrFinalLoom = new loomSort();

                        for (int i = 0; i < loomRowCount; i++)
                        {
                            if (i <= loomSortList.Count - 1)
                            {
                                arrFinalLoom = loomSortList[i];
                                arrFinalReportLOOM[i + 1, 1] = arrFinalLoom.WireCode;
                                arrFinalReportLOOM[i + 1, 2] = arrFinalLoom.length;
                                arrFinalReportLOOM[i + 1, 3] = arrFinalLoom.FromConn;
                                arrFinalReportLOOM[i + 1, 4] = arrFinalLoom.FromPin;
                                arrFinalReportLOOM[i + 1, 5] = "CRIMP";
                                arrFinalReportLOOM[i + 1, 6] = arrFinalLoom.ToConn;
                                arrFinalReportLOOM[i + 1, 7] = arrFinalLoom.ToPin;
                                arrFinalReportLOOM[i + 1, 8] = "CRIMP";
                                string[] arr = arrFinalLoom.WireCode.Split('/');
                                string cablename = arrFinalLoom.WireType.ToString();

                                if (allspecialcables.Contains(cablename))
                                {
                                    if (arr.Length > 1)
                                    {
                                        if (arr.Length > 3)
                                        {
                                            arrFinalReportLOOM[i + 1, 9] = arr[1];
                                            arrFinalReportLOOM[i + 1, 1] = arr[0] + "/" + arr[2];
                                        }
                                        else
                                        {
                                            arrFinalReportLOOM[i + 1, 9] = arrFinalLoom.WireType;
                                        }
                                    }
                                    else
                                    {
                                        continue;
                                    }
                                }
                                else
                                {
                                    arrFinalReportLOOM[i + 1, 9] = gug + "" + arr[1] + " " + arrFinalLoom.WireType;
                                }

                                arrFinalReportLOOM[i + 1, 19] = arrFinalLoom.layer;
                            }
                        }

                        int SlNo = 0;

                        int LoomInitialRow = 8;

                        int LoomSheetRowRequired = 38;
                        int NoLoomSheets = ExcelRowRules.GetNumberOfSheetsRequired(loomRowCount, LoomSheetRowRequired);

                        //implementing Merge and demege.
                        int serialNo = 1;
                        int loomMergeEndRow = 0;
                        int loomMergeStartRow;
                        int loomStartRowCount = 1;
                        int loomMergeEndRowCount = 1;
                        int listCount = 0;

                        // The loop below re-ran a full .Where(WireName == ...).Count() scan over
                        // loomSortList on every iteration to size each merge group - O(n) per
                        // iteration, effectively O(n^2) overall for large looms. Precompute each
                        // WireName's total occurrence count once instead.
                        var wireNameCounts = new Dictionary<string, int>();
                        foreach (var entry in loomSortList)
                        {
                            wireNameCounts.TryGetValue(entry.WireName, out int existing);
                            wireNameCounts[entry.WireName] = existing + 1;
                        }

                        do
                        {
                            var loomMerge = loomSortList[listCount];

                            loomStartRowCount = (listCount + 1) % LoomSheetRowRequired;

                            if (loomStartRowCount == 0)
                            {
                                loomMergeStartRow = LoomSheetRowRequired;
                            }
                            else
                            {
                                loomMergeStartRow = loomStartRowCount;
                            }

                            arrFinalReportLOOM[listCount + 1, 11] = "ME";
                            arrFinalReportLOOM[listCount + 1, 12] = (loomMergeStartRow + LoomInitialRow).ToString();
                            //arrFinalReportLOOM[loomMergeStartRow, 13] = (loomMergeEndRow+LoomInitialRow).ToString();
                            arrFinalReportLOOM[listCount + 1, 15] = arrFinalReportLOOM[listCount + 1, 9];
                            arrFinalReportLOOM[listCount + 1, 14] = serialNo;

                            int sameloomListCount = wireNameCounts[loomMerge.WireName];

                            loomMergeEndRowCount = (listCount + sameloomListCount) % LoomSheetRowRequired;

                            if (loomMergeEndRowCount == 0)
                            {
                                loomMergeEndRow = LoomSheetRowRequired;
                            }
                            else
                            {
                                loomMergeEndRow = loomMergeEndRowCount;
                            }

                            if (loomMergeStartRow > loomMergeEndRow)
                            {
                                arrFinalReportLOOM[listCount + 1, 13] = (LoomSheetRowRequired + LoomInitialRow).ToString();
                            }
                            else
                            {
                                arrFinalReportLOOM[listCount + 1, 13] = (loomMergeEndRow + LoomInitialRow).ToString();
                            }

                            listCount = listCount + sameloomListCount;

                            arrFinalReportLOOM[listCount, 11] = "ME";
                            arrFinalReportLOOM[listCount, 12] = (loomMergeStartRow + LoomInitialRow).ToString();
                            arrFinalReportLOOM[listCount, 13] = (loomMergeEndRow + LoomInitialRow).ToString();
                            arrFinalReportLOOM[listCount, 15] = arrFinalReportLOOM[listCount, 9];
                            arrFinalReportLOOM[listCount, 14] = serialNo;

                            if (loomMergeStartRow > loomMergeEndRow)
                            {
                                arrFinalReportLOOM[listCount, 12] = (LoomInitialRow + 1).ToString();
                            }

                            serialNo++;

                        } while (listCount < loomSortList.Count());


                        if (string.IsNullOrEmpty((string)arrFinalReportLOOM[loomSortList.Count, 11]))
                        {
                            arrFinalReportLOOM[loomSortList.Count, 11] = "LAST";
                            arrFinalReportLOOM[loomSortList.Count, 14] = SlNo.ToString();
                            arrFinalReportLOOM[loomSortList.Count, 12] = loomSortList.Count.ToString();
                        }

                        // modExcelInst.GenerateHALReportFormat_CableList(arrFinalReportLOOM, loomName, NoLoomSheets, LoomSheetRowRequired, ConfigurationManager.AppSettings["BundleFolder"]);
                        excel.GenerateHALReportFormat_CableList(arrFinalReportLOOM, loomName, NoLoomSheets, LoomSheetRowRequired, bundleFolderPath);
                        Logging.Info($"{loomName} Loom report Completed");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"{loomName}:incorrect cable group id");
                        Logging.Error($"{loomName}: {ex.Message}");
                    }
                });
            }
            catch (Exception ex)
            {
                Logging.Error(ex.Message);
                MessageBox.Show("ModMain ## 07" + ex.Message);
            }
        }

        // this method generates Excel reports for Megger
        public void MeggerSchedulesheet(modExcel modExcelInst, List<string> selectedLoomlistforMegger, string reportsFolderPath, List<ElectreObject> filteredElecCollection)
        {
            string[,] arrContList = null;

            var partnumbers = partnumbersList.ToList();

            // var selectedLoomlistforMegger = _customReportWindow.SelectedMeggerLoomList;

            var selectedLoomSetForMegger = new HashSet<string>(selectedLoomlistforMegger);
            var filteredData = filteredElecCollection.Where(e => selectedLoomSetForMegger.Contains(e.BundleName) && string.IsNullOrEmpty(e.NoMegger)).ToList();
           // filteredData = filteredData.Where(e => string.IsNullOrEmpty(e.NoMegger)).ToList();            

            filteredData = filteredData.OrderByDescending(obj => obj.ComponentType == "EQU") // "EQU" first
                                                  .ThenBy(obj =>
                                                      EquConnectorMap.EndsWithUnderscoredKey(obj.ConnectorName)
                                                      ? 1 : 0 // if it matches a known connector suffix → 1 (comes later), else → 0 (comes first)
                                                  )
                                                  .ThenBy(obj => obj.ConnectorName, StringComparer.OrdinalIgnoreCase)
                                                  .ThenBy(obj => obj.PinNumber, StringComparer.OrdinalIgnoreCase)
                                                  .ToList();

            bool proceedWithRemainingSheets = false;

            string templpath = Path.Combine(GlobalVar.StrtCmd, "templ");

            string fullPathMeggerFolder = Path.Combine(templpath, reportsFolderPath);

            try
            {
                // Delete the existing text files in the Megger folder
                ReportFolderCleaner.DeleteExistingFiles(fullPathMeggerFolder);
              /*  if (Directory.Exists(fullPathMeggerFolder))
                {
                    List<string> existingFiles = Directory.GetFiles(fullPathMeggerFolder, "*.txt").ToList();
                    foreach(var file in existingFiles)
                    {
                        File.Delete(file);
                    }
                }*/
            }
            catch(Exception e)
            {
                MessageBox.Show("Error in Megger Folder Path: " + e.Message);
                Logging.Error("Error in Megger Folder Path: " + e.Message);
            }

            try
            {
                // Get the LibraryPath for the Pin List
                string librayPath = Environment.GetEnvironmentVariable("ELECTRE_CUSTOMIZE") + ConfigurationManager.AppSettings["LibrayPath"];                

                string workbookPath = modExcelInst.CreateNewWorkbookforMeggerScheduler(
                    "Megger Scheduler", "Continuity Components", "Pin List", "Connection List", "Exception", "Megger", "High Megger", "Low Megger",
                    reportsFolderPath);

                if (!modExcelInst.CreateMeggerSchedulerHeader(workbookPath))
                {
                    Logging.Error("Error Component Report Header Megger Schedule");
                }

                modExcelInst.ReportWB = modExcelInst.ExcelApp.Workbooks.Open(workbookPath);

                foreach (Worksheet sheet in modExcelInst.ReportWB.Sheets)
                {
                    if (sheet.Name == "Continuity Components")
                    {
                        modExcelInst.ReportWS = sheet;
                        Dictionary<string, List<string>> matchedComponents = new Dictionary<string, List<string>>();

                        ContinuityWOBreakdown cwob = new ContinuityWOBreakdown(modExcelInst, selectedLoomlistforMegger);
                        // string[,] arrFTcwobDist = cwob.ConnectionComponentsforMegger();
                        cwob.ConnectionComponentsforMegger(filteredElecCollection);
                        foreach (var continuityComponent in arrListOfContinuityComponents)
                        {
                            if (connectorsAndPartNames.ContainsKey(continuityComponent))
                            {
                                var uniquePartNames = connectorsAndPartNames[continuityComponent].Distinct().ToList();
                                matchedComponents.Add(continuityComponent, uniquePartNames);
                                //matchedComponents.Add(continuityComponent, connectorsAndPartNames[continuityComponent].ToList());
                            }
                        }
                        var sortedMatchedComponents = matchedComponents.OrderBy(kvp => kvp.Key, StringComparer.OrdinalIgnoreCase)
                                                                        .ToDictionary(kvp => kvp.Key, kvp => kvp.Value);

                        object[,] arrCustomLoom_CC = MeggerListConverters.ToWideArray(sortedMatchedComponents);
                        modExcelInst.AppendToExcelMegger(modExcelInst.ReportWS, arrCustomLoom_CC);
                    }
                    else if (sheet.Name == "Pin List")
                    {
                        modExcelInst.ReportWS1 = sheet;
                        Microsoft.Office.Interop.Excel.Range rng = modExcelInst.ReportWS1.Columns[2];
                        rng.NumberFormat = "@";
                        var arrCustomLoom_PinList = GeneratePinListArray(partnumbers, partNumbersandConnectors, librayPath);
                        modExcelInst.AppendToExcelMegger(modExcelInst.ReportWS1, arrCustomLoom_PinList);
                    }
                    else if (sheet.Name == "Connection List")
                    {
                        ContinuityWOBreakdown cwob = new ContinuityWOBreakdown(modExcelInst, selectedLoomlistforMegger);
                        arrContList = cwob.removeRedundantsMegger(arrFTcwob);
                        modExcelInst.ReportWS2 = sheet;
                        modExcelInst.AppendToExcelMegger(modExcelInst.ReportWS2, arrContList);
                    }
                    else if (sheet.Name == "Exception")
                    {
                        string[,] arrNoMegger = MeggerListConverters.NoMeggerToArray(ElecCollectionforNoMegger); // 
                        modExcelInst.ReportWS3 = sheet;
                        modExcelInst.AppendToExcelMegger(modExcelInst.ReportWS3, arrNoMegger);

                        // Ask user whether to continue after 4 sheets
                        var result = MessageBox.Show("Do you want to generate 'Megger' text files?",
                                                     "Generate Additional Sheets?",
                                                     MessageBoxButtons.YesNo,
                                                     MessageBoxIcon.Question);

                        if (result == DialogResult.Yes)
                        {
                            proceedWithRemainingSheets = true;
                            Logging.Info("User chose to continue with 'Megger' sheets.");
                        }
                        else
                        {
                            Logging.Info($"User chose to skip 'Megger' sheets.");
                            break; // Exit the foreach loop — do not process further sheets
                        }
                    }
                    else if (proceedWithRemainingSheets && sheet.Name == "Megger" && arrContList != null)
                    {
                        lowMeggerData = MeggerListConverters.ToLowMeggerLines(arrContList); // populating lowMeggerData with Continuity data
                        // filtering the filteredData based on only the Continuity (Source and destination) components
                        // (HashSet snapshot instead of a List<string>.Contains scan per filteredData element)
                        var continuityComponentsSet = new HashSet<string>(arrListOfContinuityComponents);
                        filteredData = filteredData.Where(e => continuityComponentsSet.Contains(e.ConnectorName)).ToList();
                        MeggerSheet5(filteredData, reportsFolderPath);

                        // Add hyperlinks to Megger_*.txt
                        ExcelHyperlinks.AddTextFileLinks(sheet, fullPathMeggerFolder, "Megger_");
                    }
                    else if (proceedWithRemainingSheets && sheet.Name == "High Megger")
                    {
                        MeggerSheet6and7("High", reportsFolderPath);
                        // ✅ Add hyperlinks to HighMegger_*.txt
                        ExcelHyperlinks.AddTextFileLinks(sheet, fullPathMeggerFolder, "HighMegger_");
                    }
                    else if (proceedWithRemainingSheets && sheet.Name == "Low Megger")
                    {
                        MeggerSheet6and7("Low", reportsFolderPath);
                        // ✅ Add hyperlinks to LowMegger_*.txt
                        ExcelHyperlinks.AddTextFileLinks(sheet, fullPathMeggerFolder, "LowMegger_");
                    }
                }
                // Ensure the desired sheet is active before final save
                Worksheet firstSheet = (Worksheet)modExcelInst.ReportWB.Sheets[1];
                firstSheet.Activate();

                // Now save AFTER activating
                modExcelInst.ReportWB.Save();

            }
            catch (Exception ex)
            {
                MessageBox.Show("modMain ##007: " + ex.Message);
            }
        }

        /*  public void MeggerSchedulesheetold(modExcel modExcelInst)
          {
              string[] arrTemp1;

              string[] arrCSVrow1 = new string[1];

              List<string> selectedata = new List<string>();

              //From Data_Extraction.csv to  geting the Part numbers 
              var partnumber = partnumbersList.ToList();

              //From Data_Extraction.csv to getting the part numbers and Connector Names
              var partnumberandconnector = partnumbersandconnectorslist.ToList();

              //adding all connectors and pins purpose created the List Objacte
              List<string> allconnectorsandpins = new List<string>();

              int maxsizofconnector = 0;

              string partnumber3 = string.Empty;


              List<string> pinlistfromliabarary = new List<string>();

              try
              {
                  string librayPath = Environment.GetEnvironmentVariable("ELECTRE_CUSTOMIZE") + ConfigurationManager.AppSettings["LibrayPath"];

                  if (!File.Exists(librayPath))
                  {
                      MessageBox.Show($"File not found: {librayPath}");
                      Application.Exit();
                  }
                  int objcount = 0;
                  using (StreamReader reader = new StreamReader(librayPath))
                  {
                      ElecList.Clear();
                      while (!reader.EndOfStream)
                      {
                          objcount++;
                          arrCSVrow1[0] = reader.ReadLine() ?? string.Empty;
                          arrTemp1 = arrCSVrow1[0].Split(";");

                          selectedata.Add(arrTemp1[0]);
                          selectedata.Add(arrTemp1[4]);


                          Library Library = new Library
                          {
                              ConnectorName = arrTemp1[0],
                              Pinnumber = arrTemp1[14],
                              MaxsizeofPins = arrTemp1[4],
                              Gauge = arrTemp1[1],

                          };

                          ElecList.Add(Library);
                      }

                  }


                  int rows = 0;
                  int cols = 2;


                  object[,] arrCustomLoom = new object[rows, cols];

                  ElectreObject E17;


                  foreach (string s2 in partnumber)
                  {

                      foreach (string s in selectedata)
                      {

                          if (s == s2)
                          {
                              partnumber3 = s;
                          }

                      }
                      if (partnumber3 == s2)
                      {
                          //Maximum pins of  every Connector
                          maxsizofconnector = Convert.ToInt32(selectedata.SkipWhile(x => x != partnumber3).Skip(1).FirstOrDefault());

                          var gaug = ElecList.Where(x => x.ConnectorName == s2 && x.MaxsizeofPins == Convert.ToString(maxsizofconnector)).Select(x => x.Gauge).FirstOrDefault();

                          pinlistfromliabarary = ElecList
                             .Where(x => x.ConnectorName == s2
                              && Convert.ToString(x.MaxsizeofPins) == Convert.ToString(maxsizofconnector)
                              && Convert.ToString(x.Gauge) == Convert.ToString(gaug))
                               .Select(x => Convert.ToString(x.Pinnumber))
                              .ToList();

                          rows += pinlistfromliabarary.Count;
                          arrCustomLoom = new object[rows, 2];

                          var connector = string.Empty;
                          foreach (string s in partnumber)
                          {
                              if (partnumber3 == s)
                              {

                                  connector = partnumberandconnector.SkipWhile(x => x != partnumber3).Skip(1).FirstOrDefault();

                                  for (int i = 0; i < pinlistfromliabarary.Count; i++)
                                  {

                                      if (connector != null)
                                      {
                                          allconnectorsandpins.Add(connector);
                                      }
                                      allconnectorsandpins.Add(pinlistfromliabarary[i]);
                                  }

                              }

                          }

                          maxsizofconnector = 0;
                      }

                  }

                  //  Assign list data to the arrCustomLoom
                  int index = 0;
                  for (int i = 0; i < rows; i++)
                  {
                      for (int j = 0; j < cols; j++)
                      {
                          arrCustomLoom[i, j] = allconnectorsandpins[index];
                          index++;
                      }
                  }




                  string workbookPath = modExcelInst.CreateNewWorkbookforMeggerScheduler("Megger Scheduler", "pin List", "Connection List", "Continuity Components", "Exception",  ConfigurationManager.AppSettings["MeggerSchedulerFolder"]);

                  if (!modExcelInst.CreateMeggerSchedulerHeader(workbookPath))
                  {
                      Logging.Error($"Error Component Report Header Megger Schedule");
                  }

                  modExcelInst.ReportWB = modExcelInst.ExcelApp.Workbooks.Open(workbookPath);

                  foreach (Worksheet sheet in modExcelInst.ReportWB.Sheets)
                  {
                      if (sheet.Name == "pin List")
                      {
                          modExcelInst.ReportWS = sheet;

                          modExcelInst.AppendToExcelMegger(modExcelInst.ReportWS, arrCustomLoom);


                      }
                      else if (sheet.Name == "Connection List")
                      {
                          ElectreObject E1;

                          ContinuityWOBreakdown cwob = new ContinuityWOBreakdown(modExcelInst, _customReportWindow.SelectedMeggerLoomList);

                          modExcelInst.ReportWS1 = sheet;
                          string[,] arrFTcwobDist = cwob.ConnectionListforMegger();

                          modExcelInst.AppendToExcelMegger(modExcelInst.ReportWS1, arrFTcwobDist);



                      }
                      else if (sheet.Name == "Continuity Components")
                      {
                          modExcelInst.ReportWS2 = sheet;
                          var listOfContinuityComponents = modMain.arrListOfContinuityComponents;

                          Dictionary<string, List<string>> connectorsAndPartNames = modMain.connectorsAndPartNames;

                          Dictionary<string, List<string>> matchedComponents = new Dictionary<string, List<string>>();

                          foreach (var continuityComponent in listOfContinuityComponents)
                          {
                              if (connectorsAndPartNames.ContainsKey(continuityComponent))
                              {

                                  matchedComponents.Add(continuityComponent, connectorsAndPartNames[continuityComponent]);
                              }
                          }

                          arrCustomLoom = MeggerListConverters.ToWideArray(matchedComponents);
                          modExcelInst.AppendToExcelMegger(modExcelInst.ReportWS2, arrCustomLoom);


                      }
                      else if (sheet.Name == "Exception")
                      {

                      }


                  }


                  MeggerSheet5();
              }
              catch (Exception ex)
              {
                  MessageBox.Show("modMain ##007: " + ex.Message);
              }


          }*/

        public object[,] GeneratePinListArray(List<string> partnumbers, Dictionary<string, HashSet<string>> partNumbersandConnectors, string libraryPath)
        {
            List<string> selectedata = new List<string>();
            List<string> allconnectorsandpins = new List<string>();
            int maxsizofconnector = 0;
            string[] arrTemp1;
            string[] arrCSVrow1 = new string[1];
            List<string> pinlistfromliabarary = new List<string>();
            int cols = 2;

            try
            {
                if (!File.Exists(libraryPath))
                {
                    MessageBox.Show($"File not found: {libraryPath}");
                    return new object[0, 0];
                }

                // Read Library CSV
                using (StreamReader reader = new StreamReader(libraryPath))
                {
                    ElecList.Clear();
                    while (!reader.EndOfStream)
                    {
                        arrCSVrow1[0] = reader.ReadLine() ?? string.Empty;
                        arrTemp1 = arrCSVrow1[0].Split(";");

                        // selectedata.Add(arrTemp1[0]);
                        // selectedata.Add(arrTemp1[4]);

                        Library lib = new Library
                        {
                            PartNumber = arrTemp1[0],
                            Pinnumber = arrTemp1[14],
                            MaxsizeofPins = arrTemp1[4],
                            Gauge = arrTemp1[1]
                        };

                        ElecList.Add(lib);
                    }
                }

                // Build once: group the library entries by PartNumber, and index ElecCollection's
                // Gauge by ConnectorName (first occurrence, matching the original FirstOrDefault),
                // plus a HashSet snapshot of arrListOfContinuityComponents. Avoids each
                // (partNumber, connectorName) pair below rescanning the whole ElecList/
                // ElecCollection/arrListOfContinuityComponents, which is what made this loop
                // O(partNumbers * connectorNames * ElecList.Count) in practice.
                var elecListByPartNumber = new Dictionary<string, List<Library>>();
                foreach (var lib in ElecList)
                {
                    if (!elecListByPartNumber.TryGetValue(lib.PartNumber, out var list))
                    {
                        list = new List<Library>();
                        elecListByPartNumber[lib.PartNumber] = list;
                    }
                    list.Add(lib);
                }

                var gaugeByConnectorName = new Dictionary<string, string>();
                foreach (var e in ElecCollection)
                {
                    if (!gaugeByConnectorName.ContainsKey(e.ConnectorName))
                    {
                        gaugeByConnectorName[e.ConnectorName] = e.Gauge;
                    }
                }

                var continuityComponentsSet = new HashSet<string>(arrListOfContinuityComponents);

                // The pins a connector gets depend only on (part number, gauge of that connector in the data), and many
                // connectors share the same pair - so the library lookup and the pin-range expansion are done once per pair.
                var expandedPinsByPartAndGauge = new Dictionary<(string PartNumber, string Gauge), List<string>>();

                foreach (string partNumber in partnumbers)
                {
                    if (!elecListByPartNumber.TryGetValue(partNumber, out var libEntriesForPart))
                        continue;

                    var connectorNames = partNumbersandConnectors[partNumber];

                    foreach (var connectorName in connectorNames)
                    {
                        // Match only connectors listed in continuity components
                        if (string.IsNullOrEmpty(connectorName) || !continuityComponentsSet.Contains(connectorName))
                            continue;

                        // Retrive the Gauge from data extraction based on the matched ConnectorName
                        var connectorGauge_dataextraction = gaugeByConnectorName.TryGetValue(connectorName, out var gaugeVal) ? gaugeVal : null;

                        if (!expandedPinsByPartAndGauge.TryGetValue((partNumber, connectorGauge_dataextraction), out var expandedPins))
                        {
                            // Retrive the gauge of the Part Number from the Library
                            var gauge_fromLibrary = libEntriesForPart.Where(x => x.Gauge == connectorGauge_dataextraction)
                                                             .Select(x => x.Gauge)
                                                             .FirstOrDefault();

                            // If both dataextraction and library Gauges are matched then retrive the maxSizeofConnector
                            maxsizofconnector = libEntriesForPart.Where(e => e.Gauge == gauge_fromLibrary)
                                                        .Select(e => e.MaxsizeofPins)
                                                        .Select(p => Convert.ToInt32(p))
                                                        .FirstOrDefault();

                            // Retrive the pins List of the matched Part number, MaxsizeofPins and Gauge
                            pinlistfromliabarary = libEntriesForPart
                                .Where(x => x.MaxsizeofPins == maxsizofconnector.ToString()
                                    && x.Gauge == gauge_fromLibrary)
                                .Select(x => x.Pinnumber)
                                .ToList();

                            expandedPins = LibraryPinExpander.Expand(pinlistfromliabarary);
                            expandedPinsByPartAndGauge[(partNumber, connectorGauge_dataextraction)] = expandedPins;
                        }

                        // Adding every pin of this connector into allconnectorsandpins (connector, pin, connector, pin ...)
                        foreach (string pin in expandedPins)
                        {
                            allconnectorsandpins.Add(connectorName);
                            allconnectorsandpins.Add(pin);
                        }

                        maxsizofconnector = 0;
                    }
                }

                // Create 2D array for Pin list
                int Rows = allconnectorsandpins.Count / 2;
                object[,] arrCustomLoom_PinList = new object[Rows, cols];
                int index = 0;
                for (int i = 0; i < Rows; i++)
                {
                    for (int j = 0; j < cols; j++)
                    {
                        arrCustomLoom_PinList[i, j] = allconnectorsandpins[index++];
                    }
                }

                var dataList = new List<(string Connector, string Pin)>();

                // Step 1: Convert 2D array to list of tuples
                for (int i = 0; i < Rows; i++)
                {
                    dataList.Add((
                        arrCustomLoom_PinList[i, 0]?.ToString() ?? string.Empty,
                        arrCustomLoom_PinList[i, 1]?.ToString() ?? string.Empty
                    ));
                }

                // Step 2: Sort list (by Connector, then Pin) — Simple ordinal sorting
                var sortedList = dataList
                    .OrderBy(x => x.Connector, StringComparer.OrdinalIgnoreCase)
                    .ToList();

                // Step 3: Write back into 2D array
                for (int i = 0; i < Rows; i++)
                {
                    arrCustomLoom_PinList[i, 0] = sortedList[i].Connector;
                    arrCustomLoom_PinList[i, 1] = sortedList[i].Pin;
                }


                // Populate the List for High Megger data from the Pin List data  from Library
                listConnectorPinsLibrary = new List<(string ConnectorName, string PinNumber)>();

                for (int i = 0; i < allconnectorsandpins.Count - 1; i += 2) 
                {
                    string connectorName = allconnectorsandpins[i];
                    string pinNumber = allconnectorsandpins[i + 1];

                    if (!string.IsNullOrEmpty(connectorName) && !string.IsNullOrEmpty(pinNumber))
                    {
                        listConnectorPinsLibrary.Add((connectorName, pinNumber));
                    }
                }


                return arrCustomLoom_PinList;
            }
            catch (Exception ex)
            {
                MessageBox.Show("GeneratePinListArray Error: " + ex.Message);
                return new object[0, 0];
            }
        }

        // Streams Megger_{n}.txt (lowMeggerData lines first, then the High Megger pairs).
        // All data work lives in MeggerReportWriter / the Megger engine; this method only keeps the
        // UI concerns (folder, logging, message box).
        public void MeggerSheet5(List<ElectreObject> filteredData, string reportsFolderPath)
        {
            try
            {
                string templpath = Path.Combine(GlobalVar.StrtCmd, "templ");
                string folderpath = Path.Combine(templpath, reportsFolderPath);

                if (!Directory.Exists(folderpath))
                    Directory.CreateDirectory(folderpath);

                MeggerReportWriter.Reset();
                highMeggerData = new ConcurrentBag<string>();

                if (filteredData.Count == 0)
                {
                    return;
                }

                var files = MeggerReportWriter.WriteCombinedSheets(
                    folderpath, filteredData, lowMeggerData, listConnectorPinsLibrary,
                    (totalLines, totalSheets) =>
                    {
                        Logging.Info($"Total Megger lines to write: {totalLines}");
                        Logging.Info($"Total Megger sheets to create: {totalSheets}");
                    });

                foreach (var file in files)
                {
                    Logging.Info($"{Path.GetFileName(file)} created successfully");
                }

                Logging.Info("Megger sheet generation completed successfully.");
            }
            catch (Exception ex)
            {
                var real = (ex as AggregateException)?.Flatten().InnerException ?? ex;
                MessageBox.Show($" Error Megger sheet: {real.Message}");
                Logging.Error($"Megger Sheet 5: Error occured : {real.Message}");
            }
        }

        // Writes HighMegger_{n}.txt (High Megger pairs only) or {type}Megger_{n}.txt (lowMeggerData only).
        public void MeggerSheet6and7(string meggerType, string reportsFolderPath)
        {
            try
            {
                string templpath = Path.Combine(GlobalVar.StrtCmd, "templ");
                string folderpath = Path.Combine(templpath, reportsFolderPath);

                if (!Directory.Exists(folderpath))
                {
                    Directory.CreateDirectory(folderpath);
                }

                var files = MeggerReportWriter.WriteTypeSheets(meggerType, folderpath, lowMeggerData);

                foreach (var file in files)
                {
                    Logging.Info($"{Path.GetFileName(file)} created successfully");
                }

                Logging.Info($"{meggerType} Megger sheet generation completed successfully.");
            }
            catch (Exception ex)
            {
                var real = (ex as AggregateException)?.Flatten().InnerException ?? ex;
                MessageBox.Show($"Error {meggerType}: " + real.Message);
                Logging.Error($"Error occurred in {meggerType} Megger sheet: {real.Message}");
            }
        }

        // Commented on July 17 2025 before implementing splitting the files 
        /* public void MeggerSheet6and7(string meggerType)
         {
             try
             {
                 string templpath = Path.Combine(GlobalVar.StrtCmd, "templ");
                 string folderpath = Path.Combine(templpath, "MeggerSchedulerReports");

                 if (!Directory.Exists(folderpath))
                 {
                     Directory.CreateDirectory(folderpath);
                 }

                 string filePath = Path.Combine(folderpath, $"{meggerType}Megger.txt");
                 IEnumerable<string> outputLines = meggerType == "High" ? highMeggerData : lowMeggerData;

                 using (var writer = new StreamWriter(filePath, false, Encoding.UTF8, bufferSize: 65536)) // 64 KB buffer
                 {
                     writer.WriteLine("From;TO");
                     writer.WriteLine("FromConnector;FromPin;ToConnector;ToPin;Megger");

                     foreach (var line in outputLines)
                     {
                         writer.WriteLine(line);
                     }
                 }

                 Logging.Info($"{meggerType} Megger sheet completed successfully.");
             }
             catch (Exception ex)
             {
                 MessageBox.Show($"Error {meggerType}: " + ex.Message);
                 Logging.Error($"Error occured {meggerType} : {ex.Message}");
             }
         }*/

        #region Megger Sheet 5, 6 and 7 old methods
        // Commented on June 27, 2025
        /*   public void MeggerSheet5(List<ElectreObject> filteredData)
        {
            try
            {
                string templpath = Path.Combine(GlobalVar.StrtCmd, "templ");
                string folderpath = Path.Combine(templpath, "MeggerSchedulerReports");
                string finalFilePath = Path.Combine(folderpath, "Megger.txt");

                if (!Directory.Exists(folderpath))
                    Directory.CreateDirectory(folderpath);

                highMeggerData = new ConcurrentBag<string>();
                // lowMeggerData = new ConcurrentBag<string>();

                var tempFiles = new ConcurrentBag<string>();

                int count = filteredData.Count;
                var redundancySet = new ConcurrentDictionary<string, byte>(Environment.ProcessorCount * 2, count * 2);
                //int partitionSize = 1000;
                int partitionSize = Math.Max(200, count / Environment.ProcessorCount);
                var rangePartitioner = Partitioner.Create(0, count, partitionSize);

                Parallel.ForEach(rangePartitioner, new ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount }, range =>
                {
                    var localBuilder = new StringBuilder(100_000);

                    for (int i = range.Item1; i < range.Item2; i++)
                    {
                        var E1 = filteredData[i];
                        //if (string.IsNullOrEmpty(E1.WireNumber)) continue;

                        for (int j = 0; j < count; j++)
                        {
                            if (i == j) continue;
                            var n1 = filteredData[j];

                            // Ensure symmetric key to avoid A-B and B-A duplicates
                            string partA = $"{E1.ConnectorName}#{E1.PinNumber}";
                            string partB = $"{n1.ConnectorName}#{n1.PinNumber}";
                            string key = string.Compare(partA, partB, StringComparison.OrdinalIgnoreCase) < 0
                                ? $"{partA}##{partB}"
                                : $"{partB}##{partA}";

                            if (!redundancySet.TryAdd(key, 0))
                                continue;

                            string status = (string.Equals(n1.WireNumber, E1.WireNumber, StringComparison.OrdinalIgnoreCase) &&
                                             string.Equals(n1.SubNet, E1.SubNet, StringComparison.OrdinalIgnoreCase))
                                             ? "Low Megger"
                                             : "High Megger";

                            string resultLine = string.Empty;

                            if (status == "High Megger")
                            {
                                resultLine = $"{E1.ConnectorName};{E1.PinNumber};{n1.ConnectorName};{n1.PinNumber};{status}";
                                //highMeggerData.Add(resultLine);
                                localBuilder.AppendLine(resultLine);
                            }
                        }
                    }

                    string tempFile = Path.Combine(folderpath, $"temp_{Guid.NewGuid()}.txt");
                    File.WriteAllText(tempFile, localBuilder.ToString());
                    tempFiles.Add(tempFile);
                });

                // High Megger data for Pin List Library Pins Extra data

                var highMeggerBuilder = new StringBuilder(200_000);
                var builderLock = new object(); // for thread-safe string building

                // Partition the source list for parallel processing
                var rangePartitioner1 = Partitioner.Create(0, listConnectorPinsLibrary.Count);

                Parallel.ForEach(rangePartitioner1, new ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount }, range =>
                {
                    var localBuilder1 = new StringBuilder(50_000);

                    for (int idx = range.Item1; idx < range.Item2; idx++)
                    {
                        var source = listConnectorPinsLibrary[idx];
                        string srcConnector = source.ConnectorName;
                        string srcPin = source.PinNumber;

                        // Loop through filteredData
                        foreach (var target in filteredData)
                        {
                            string tgtConnector = target.ConnectorName;
                            string tgtPin = target.PinNumber;

                            if (string.Equals(srcConnector, tgtConnector, StringComparison.OrdinalIgnoreCase) &&
                                string.Equals(srcPin, tgtPin, StringComparison.OrdinalIgnoreCase))
                                continue;

                            // Create symmetric key
                            string keyA = $"{srcConnector}#{srcPin}";
                            string keyB = $"{tgtConnector}#{tgtPin}";
                            string pairKey = string.Compare(keyA, keyB, StringComparison.OrdinalIgnoreCase) < 0
                                ? $"{keyA}##{keyB}"
                                : $"{keyB}##{keyA}";

                            if (!redundancySet.TryAdd(pairKey, 0))
                                continue;

                            string line = $"{srcConnector};{srcPin};{tgtConnector};{tgtPin};High Megger";
                            localBuilder1.AppendLine(line);
                            //highMeggerData.Add(line);
                        }

                        // Loop through pin list itself
                        for (int j = 0; j < listConnectorPinsLibrary.Count; j++)
                        {
                            if (j == idx) continue;
                            var target = listConnectorPinsLibrary[j];
                            string tgtConnector = target.ConnectorName;
                            string tgtPin = target.PinNumber;

                            // Create symmetric key
                            string keyA = $"{srcConnector}#{srcPin}";
                            string keyB = $"{tgtConnector}#{tgtPin}";
                            string pairKey = string.Compare(keyA, keyB, StringComparison.OrdinalIgnoreCase) < 0
                                ? $"{keyA}##{keyB}"
                                : $"{keyB}##{keyA}";

                            if (!redundancySet.TryAdd(pairKey, 0))
                                continue;

                            string line = $"{srcConnector};{srcPin};{tgtConnector};{tgtPin};High Megger";
                            localBuilder1.AppendLine(line);
                            // highMeggerData.Add(line);
                        }
                    }
                    // Safely merge localBuilder into main builder
                    lock (builderLock)
                    {
                        highMeggerBuilder.Append(localBuilder1);
                    }
                });

                var highMeggerBuilder = new StringBuilder(200_000);

                foreach (var source in listConnectorPinsLibrary)
                {
                    string sourceConnector = source.ConnectorName;
                    string sourcePin = source.PinNumber;

                    foreach (var target in filteredData)
                    {
                        string targetConnector = target.ConnectorName;
                        string targetPin = target.PinNumber;

                        // Skip if same connector & pin (self-connection)
                        if (string.Equals(sourceConnector, targetConnector, StringComparison.OrdinalIgnoreCase) &&
                            string.Equals(sourcePin, targetPin, StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }

                        string partA = $"{sourceConnector}#{sourcePin}";
                        string partB = $"{targetConnector}#{targetPin}";
                        string key = string.Compare(partA, partB, StringComparison.OrdinalIgnoreCase) < 0
                            ? $"{partA}##{partB}"
                            : $"{partB}##{partA}";

                        if (!redundancySet.TryAdd(key, 0))
                            continue;

                        string line = $"{sourceConnector};{sourcePin};{target.ConnectorName};{target.PinNumber};High Megger";
                        highMeggerBuilder.AppendLine(line);
                        highMeggerData.Add(line);
                    }

                    foreach (var target in listConnectorPinsLibrary)
                    {
                        string targetConnector = target.ConnectorName;
                        string targetPin = target.PinNumber;

                        // Skip self-connection
                        if (string.Equals(sourceConnector, targetConnector, StringComparison.OrdinalIgnoreCase) &&
                            string.Equals(sourcePin, targetPin, StringComparison.OrdinalIgnoreCase))
                            continue;

                        // Create symmetric key to avoid duplicates (A-B and B-A)
                        string partA = $"{sourceConnector}#{sourcePin}";
                        string partB = $"{targetConnector}#{targetPin}";
                        string key = string.Compare(partA, partB, StringComparison.OrdinalIgnoreCase) < 0
                            ? $"{partA}##{partB}"
                            : $"{partB}##{partA}";

                        if (!redundancySet.TryAdd(key, 0))
                            continue;

                        string line = $"{sourceConnector};{sourcePin};{targetConnector};{targetPin};High Megger";
                        highMeggerBuilder.AppendLine(line);
                        highMeggerData.Add(line);
                    }
                }

                var finalBuilder = new StringBuilder(3_00_000);
                finalBuilder.AppendLine("From;TO");
                finalBuilder.AppendLine("FromConnector;FromPin;ToConnector;ToPin;Megger");

                // First adding the Low Megger data into Megger sheet
                finalBuilder.Append(string.Join(Environment.NewLine, lowMeggerData));
                finalBuilder.AppendLine();
       

                // Adding all the Temp files data into Final Megger sheet in a single shot
                foreach (string tempFile in tempFiles)
                {
                    finalBuilder.Append(File.ReadAllText(tempFile));
                    File.Delete(tempFile); // Delete the temp file after using
                }

                finalBuilder.Append(highMeggerBuilder.ToString()); // adding the extra pins data from Pin List Library 

                File.WriteAllText(finalFilePath, finalBuilder.ToString());

                Logging.Info("Megger sheet generation completed successfully.");
            }
            catch (Exception ex)
            {
                MessageBox.Show($" Error Megger sheet: {ex.Message}");
                Logging.Error($"Error occured : {ex.Message}");
            }
        }*/

        // advanced without partioner
        /* public void MeggerSheet5_old(List<ElectreObject> filteredData)
         {
             try
             {
                 string templpath = Path.Combine(GlobalVar.StrtCmd, "templ");
                 string folderpath = Path.Combine(templpath, "MeggerSchedulerReports");
                 string filePath = Path.Combine(folderpath, "Megger.txt");

                 if (!Directory.Exists(folderpath))
                 {
                     Directory.CreateDirectory(folderpath);
                 }

                 int count = filteredData.Count;

                 // Clear old data
                 highMeggerData = new ConcurrentBag<string>();
                 lowMeggerData = new ConcurrentBag<string>();

                 // Thread-safe collections
                 var outputLines = new ConcurrentBag<string>();
                 var redundancyCheck = new ConcurrentDictionary<string, byte>();

                 // Parallel outer loop
                 Parallel.For(0, count, i =>
                 {
                     var E1 = filteredData[i];
                     if (string.IsNullOrEmpty(E1.WireNumber))
                         return;

                     for (int j = 0; j < count; j++)
                     {
                         if (i == j) continue;

                         var n1 = filteredData[j];

                         // Redundancy check logic
                         string connectLine = $"{n1.ConnectorName}#{n1.PinNumber}##{E1.ConnectorName}#{E1.PinNumber}";
                         string presentLine = $"{E1.ConnectorName}#{E1.PinNumber}##{n1.ConnectorName}#{n1.PinNumber}";

                         if (redundancyCheck.ContainsKey(presentLine))
                             continue;

                         string meggerStatus = (string.Equals(n1.WireNumber, E1.WireNumber, StringComparison.OrdinalIgnoreCase) &&
                                                string.Equals(n1.SubNet, E1.SubNet, StringComparison.OrdinalIgnoreCase))
                                                ? "Low Megger"
                                                : "High Megger";
                         string resultLine = $"{E1.ConnectorName};{E1.PinNumber};{n1.ConnectorName};{n1.PinNumber};{meggerStatus}";
                         if (meggerStatus == "High Megger")
                         {
                             highMeggerData.Add(resultLine);
                         }
                         if(meggerStatus == "Low Megger")
                         {
                             lowMeggerData.Add(resultLine);
                         }

                         // Add to result + redundancy set (thread-safe)
                         outputLines.Add(resultLine);
                         redundancyCheck.TryAdd(connectLine, 0);
                     }
                 });

                 // Write to file
                 var sb = new StringBuilder();
                 sb.AppendLine("From;TO");
                 sb.AppendLine("FromConnector;FromPin;ToConnector;ToPin;Megger");

                 foreach (var line in outputLines)
                 {
                     sb.AppendLine(line);
                 }

                 File.WriteAllText(filePath, sb.ToString());

                 Logging.Info("Multi-threaded Megger sheet completed successfully.");
             }
             catch (Exception ex)
             {
                 MessageBox.Show("modMain #008: " + ex.Message);
             }
         }*/

        // advanced with partioner
        /* public void MeggerSheet5Advanced(List<ElectreObject> filteredData)
         {
             try
             {
                 string templpath = Path.Combine(GlobalVar.StrtCmd, "templ");
                 string folderpath = Path.Combine(templpath, "MeggerSchedulerReports");
                 string finalFilePath = Path.Combine(folderpath, "Megger.txt");

                 if (!Directory.Exists(folderpath))
                     Directory.CreateDirectory(folderpath);

                 int count = filteredData.Count;
                 var redundancySet = new ConcurrentDictionary<string, byte>(Environment.ProcessorCount * 2, count * 2);

                 // Clear old results
                 highMeggerData = new ConcurrentBag<string>();
                 lowMeggerData = new ConcurrentBag<string>();

                 // Temp files for each partition
                 var tempFiles = new ConcurrentBag<string>();

                 int partitionSize = 1000;
                 var rangePartitioner = Partitioner.Create(0, count, partitionSize);

                 Parallel.ForEach(rangePartitioner, new ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount }, range =>
                 {
                     var localBuilder = new StringBuilder(100_000);
                     for (int i = range.Item1; i < range.Item2; i++)
                     {
                         var E1 = filteredData[i];
                         if (string.IsNullOrEmpty(E1.WireNumber)) continue;

                         for (int j = i + 1; j < count; j++)
                         {
                             var n1 = filteredData[j];

                             // Redundancy check
                             string forward = $"{E1.ConnectorName}#{E1.PinNumber}##{n1.ConnectorName}#{n1.PinNumber}";
                             string reverse = $"{n1.ConnectorName}#{n1.PinNumber}##{E1.ConnectorName}#{E1.PinNumber}";

                             if (!redundancySet.TryAdd(forward, 0))
                                 continue;

                             string status = (string.Equals(n1.WireNumber, E1.WireNumber, StringComparison.OrdinalIgnoreCase) &&
                                              string.Equals(n1.SubNet, E1.SubNet, StringComparison.OrdinalIgnoreCase))
                                              ? "Low Megger"
                                              : "High Megger";

                             string resultLine = $"{E1.ConnectorName};{E1.PinNumber};{n1.ConnectorName};{n1.PinNumber};{status}";

                             // Store in respective collections
                             if (status == "High Megger")
                                 highMeggerData.Add(resultLine);
                             else
                                 lowMeggerData.Add(resultLine);

                             // Append to local buffer
                             localBuilder.AppendLine(resultLine);
                         }
                     }

                     // Write to partition temp file
                     string tempFile = Path.Combine(folderpath, $"temp_{Guid.NewGuid()}.txt");
                     File.WriteAllText(tempFile, localBuilder.ToString());
                     tempFiles.Add(tempFile);
                 });

                 // Merge all temp files into final report
                 var finalBuilder = new StringBuilder(500_000);
                 finalBuilder.AppendLine("From;TO");
                 finalBuilder.AppendLine("FromConnector;FromPin;ToConnector;ToPin;Megger");

                 foreach (string tempFile in tempFiles)
                 {
                     finalBuilder.Append(File.ReadAllText(tempFile));
                     File.Delete(tempFile);
                 }

                 File.WriteAllText(finalFilePath, finalBuilder.ToString());

                 Logging.Info("✅ Optimized Megger sheet generation completed successfully.");
             }
             catch (Exception ex)
             {
                 MessageBox.Show("modMain #008: " + ex.Message);
             }
         }*/

        /* public void MeggerSheet6and7_oldAdvanced(string meggerType)
         {
             try
             {
                 string templpath = Path.Combine(GlobalVar.StrtCmd, "templ");
                 string folderpath = Path.Combine(templpath, "MeggerSchedulerReports");

                 if (!Directory.Exists(folderpath))
                 {
                     Directory.CreateDirectory(folderpath);
                 }

                 string filePath = Path.Combine(folderpath, $"{meggerType}Megger.txt");

                 IEnumerable<string> outputLines  = meggerType == "High" ? highMeggerData : lowMeggerData;
                 // Write to file
                 var sb = new StringBuilder();
                 sb.AppendLine("From;TO");
                 sb.AppendLine("FromConnector;FromPin;ToConnector;ToPin;Megger");

                 foreach (var line in outputLines)
                 {
                     sb.AppendLine(line);
                 }
                 File.WriteAllText(filePath, sb.ToString());

                 Logging.Info($"{meggerType} Megger sheet completed successfully.");
             }
             catch (Exception ex)
             {
                 MessageBox.Show("modMain #008: " + ex.Message);
             }
         }*/

        //old sheet5 method commented on June 4 due to lack of performace
        /*public void MeggerSheet5(List<ElectreObject> filteredData)
        {
            string templpath;
            string folderpath;
            string filePath;
            templpath = Path.Combine(GlobalVar.StrtCmd, "templ");
            folderpath = Path.Combine(templpath, "MeggerSchedulerReports");
            filePath = Path.Combine(folderpath, "Megger.txt");

            if (!Directory.Exists(folderpath))
            {
                Directory.CreateDirectory(folderpath);
            }

            try
            {
                ElectreObject E1;
                ElectreObject n1;
                int Rowcount = 0;
                listmegger.Clear();

                for (int i = 0; i < filteredData.Count; i++)
                {
                    E1 = filteredData[i];
                    if (!string.IsNullOrEmpty(E1.WireNumber))
                    {
                        for (int j = 0; j < filteredData.Count; j++)
                        {
                            n1 = filteredData[j];
                            if (j != i)
                            {
                                if (string.Equals(n1.WireNumber, E1.WireNumber, StringComparison.OrdinalIgnoreCase) &&
                                string.Equals(n1.SubNet, E1.SubNet, StringComparison.OrdinalIgnoreCase))
                                {
                                    listmegger.Add(new MeggerData
                                    {
                                        FromConnector = E1.ConnectorName,
                                        FromPin = E1.PinNumber,
                                        ToConnector = n1.ConnectorName,
                                        ToPin = n1.PinNumber,
                                        MeggerStatus = "Low Megger"
                                    });


                                }
                                else
                                {
                                    listmegger.Add(new MeggerData
                                    {
                                        FromConnector = E1.ConnectorName,
                                        FromPin = E1.PinNumber,
                                        ToConnector = n1.ConnectorName,
                                        ToPin = n1.PinNumber,
                                        MeggerStatus = "High Megger"
                                    });


                                }
                                Rowcount++;
                            }

                        }

                    }

                }
                //  Assign list data to 2d array arrmeger
                int rows = listmegger.Count;

                arrmegger = new string[rows, 5];

                for (int i = 0; i < rows; i++)
                {
                    arrmegger[i, 0] = listmegger[i].FromConnector;
                    arrmegger[i, 1] = listmegger[i].FromPin;
                    arrmegger[i, 2] = listmegger[i].ToConnector;
                    arrmegger[i, 3] = listmegger[i].ToPin;
                    arrmegger[i, 4] = listmegger[i].MeggerStatus;
                }
                string[,] arrFTcwobDist;
                arrFTcwobDist = new string[arrmegger.GetLength(0), 5];
                List<string> rededuntCheck = new List<string>();
                int arrFTcwobDistCount = 0;

                for (int i = 0; i < arrmegger.GetLength(0); i++)
                {
                    string connectLine = $"{arrmegger[i, 2]}#{arrmegger[i, 3]}##{arrmegger[i, 0]}#{arrmegger[i, 1]}";
                    string presentLine = $"{arrmegger[i, 0]}#{arrmegger[i, 1]}##{arrmegger[i, 2]}#{arrmegger[i, 3]}";

                    if (!rededuntCheck.Contains(presentLine))//to skip the redunt line.
                    {

                        arrFTcwobDist[arrFTcwobDistCount, 0] = arrmegger[i, 0];
                        arrFTcwobDist[arrFTcwobDistCount, 1] = arrmegger[i, 1];
                        arrFTcwobDist[arrFTcwobDistCount, 2] = arrmegger[i, 2];
                        arrFTcwobDist[arrFTcwobDistCount, 3] = arrmegger[i, 3];
                        arrFTcwobDist[arrFTcwobDistCount, 4] = arrmegger[i, 4];

                        arrFTcwobDistCount++;
                        rededuntCheck.Add(connectLine);
                    }
                }
                using (StreamWriter writer = new StreamWriter(filePath, false))
                {
                    //  Header
                    writer.WriteLine("From;TO\n");
                    writer.WriteLine("FromConnector;FromPin;ToConnector;ToPin;Megger");

                    int validRowCount = 0;

                    for (int i = 0; i < arrFTcwobDist.GetLength(0); i++)  // Loop Rows
                    {
                        bool isValid = false;

                        for (int j = 0; j < arrFTcwobDist.GetLength(1); j++)  // Loop Columns
                        {
                            if (!string.IsNullOrEmpty(arrFTcwobDist[i, j]))
                            {
                                isValid = true;  // At least one column has data
                                break;
                            }
                        }

                        if (isValid)
                        {
                            validRowCount++;
                        }
                    }


                    for (int i = 0; i < validRowCount; i++)  // Rows
                    {
                        string line = "";

                        for (int j = 0; j < arrFTcwobDist.GetLength(1); j++)  // Columns
                        {
                            line += arrFTcwobDist[i, j].ToString();

                            if (j < arrFTcwobDist.GetLength(1) - 1)
                                line += ";";  // Separator
                        }

                        writer.WriteLine(line);
                    }
                }


                Logging.Info("Megger sheet4 completed");
            }

            catch (Exception ex)
            {
                MessageBox.Show("modMain #008: " + ex.Message);
            }


        }*/

        //old sheet5 method commented on June 4 due to lack of performace
        /*public void MeggerSheet5_singlethread(List<ElectreObject> filteredData)
        {
            try
            {
                string templpath = Path.Combine(GlobalVar.StrtCmd, "templ");
                string folderpath = Path.Combine(templpath, "MeggerSchedulerReports");
                string filePath = Path.Combine(folderpath, "Megger.txt");

                if (!Directory.Exists(folderpath))
                {
                    Directory.CreateDirectory(folderpath);
                }

                StringBuilder sb = new StringBuilder();
                HashSet<string> redundancyCheck = new HashSet<string>();

                sb.AppendLine("From;TO");
                sb.AppendLine("FromConnector;FromPin;ToConnector;ToPin;Megger");

                int count = filteredData.Count;

                for (int i = 0; i < count; i++)
                {
                    var E1 = filteredData[i];

                    if (string.IsNullOrEmpty(E1.WireNumber))
                        continue;

                    for (int j = 0; j < count; j++)
                    {
                        if (i == j) continue;

                        var n1 = filteredData[j];

                        string connectLine = $"{n1.ConnectorName}#{n1.PinNumber}##{E1.ConnectorName}#{E1.PinNumber}";
                        string presentLine = $"{E1.ConnectorName}#{E1.PinNumber}##{n1.ConnectorName}#{n1.PinNumber}";

                        if (redundancyCheck.Contains(presentLine))
                            continue;

                        string meggerStatus = (string.Equals(n1.WireNumber, E1.WireNumber, StringComparison.OrdinalIgnoreCase) &&
                                               string.Equals(n1.SubNet, E1.SubNet, StringComparison.OrdinalIgnoreCase))
                                                ? "Low Megger"
                                                : "High Megger";

                        sb.AppendLine($"{E1.ConnectorName};{E1.PinNumber};{n1.ConnectorName};{n1.PinNumber};{meggerStatus}");
                        redundancyCheck.Add(connectLine);
                    }
                }

                File.WriteAllText(filePath, sb.ToString());
                Logging.Info("Megger sheet completed successfully.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("modMain #008: " + ex.Message);
            }
        }*/

        //old sheet6and7 method commented on June 4 due to lack of performace
        /* public void MeggerSheet6and7(List<ElectreObject> filteredData, string meggerType)
         {
             string filePath = string.Empty;
             string templpath = Path.Combine(GlobalVar.StrtCmd, "templ");
             string folderpath = Path.Combine(templpath, "MeggerSchedulerReports");
             List<MeggerData> filteredMeggerData = new List<MeggerData>();
             if (!Directory.Exists(folderpath))
             {
                 Directory.CreateDirectory(folderpath);
             }

             filePath = Path.Combine(folderpath, $"{meggerType}Megger.txt");

             if (meggerType == "High")
             {
                 filteredMeggerData = listmegger.Where(e => e.MeggerStatus == "High Megger").ToList();
             }
             else if (meggerType == "Low")
             {
                 filteredMeggerData = listmegger.Where(e => e.MeggerStatus == "Low Megger").ToList();
             }

             try
             {
                 int rows = filteredMeggerData.Count;

                 arrmegger = new string[rows, 5];

                 for (int i = 0; i < rows; i++)
                 {
                     arrmegger[i, 0] = filteredMeggerData[i].FromConnector;
                     arrmegger[i, 1] = filteredMeggerData[i].FromPin;
                     arrmegger[i, 2] = filteredMeggerData[i].ToConnector;
                     arrmegger[i, 3] = filteredMeggerData[i].ToPin;
                     arrmegger[i, 4] = filteredMeggerData[i].MeggerStatus;
                 }
                 string[,] arrMeggerdata;
                 arrMeggerdata = new string[arrmegger.GetLength(0), 5];
                 List<string> rededuntCheck = new List<string>();
                 int arrFTcwobDistCount = 0;

                 for (int i = 0; i < arrmegger.GetLength(0); i++)
                 {
                     string connectLine = $"{arrmegger[i, 2]}#{arrmegger[i, 3]}##{arrmegger[i, 0]}#{arrmegger[i, 1]}";
                     string presentLine = $"{arrmegger[i, 0]}#{arrmegger[i, 1]}##{arrmegger[i, 2]}#{arrmegger[i, 3]}";

                     if (!rededuntCheck.Contains(presentLine))//to skip the redunt line.
                     {

                         arrMeggerdata[arrFTcwobDistCount, 0] = arrmegger[i, 0];
                         arrMeggerdata[arrFTcwobDistCount, 1] = arrmegger[i, 1];
                         arrMeggerdata[arrFTcwobDistCount, 2] = arrmegger[i, 2];
                         arrMeggerdata[arrFTcwobDistCount, 3] = arrmegger[i, 3];
                         arrMeggerdata[arrFTcwobDistCount, 4] = arrmegger[i, 4];

                         arrFTcwobDistCount++;
                         rededuntCheck.Add(connectLine);
                     }
                 }
                 using (StreamWriter writer = new StreamWriter(filePath, false))
                 {
                     //  Header
                     writer.WriteLine("From;TO\n");
                     writer.WriteLine("FromConnector;FromPin;ToConnector;ToPin;Megger");

                     int validRowCount = 0;

                     for (int i = 0; i < arrMeggerdata.GetLength(0); i++)  // Loop Rows
                     {
                         bool isValid = false;

                         for (int j = 0; j < arrMeggerdata.GetLength(1); j++)  // Loop Columns
                         {
                             if (!string.IsNullOrEmpty(arrMeggerdata[i, j]))
                             {
                                 isValid = true;  // At least one column has data
                                 break;
                             }
                         }

                         if (isValid)
                         {
                             validRowCount++;
                         }
                     }


                     for (int i = 0; i < validRowCount; i++)  // Rows
                     {
                         string line = "";

                         for (int j = 0; j < arrMeggerdata.GetLength(1); j++)  // Columns
                         {
                             line += arrMeggerdata[i, j].ToString();

                             if (j < arrMeggerdata.GetLength(1) - 1)
                                 line += ";";  // Separator
                         }

                         writer.WriteLine(line);
                     }
                 }


                 Logging.Info("Megger sheet completed");
             }

             catch (Exception ex)
             {
                 MessageBox.Show("modMain #008: " + ex.Message);
             }


         }*/
        #endregion
               
        public void CableList_Report_Preprocessing(object[,] iarr, string reportName, modExcel modExcelInst)
        {
            try
            {
                int maxRows = ExcelRowRules.MaxRowinArray(iarr);
                int finalRowCount = 0;

                const int initialRowOffset = 1;  // Excel is 1-indexed
                const int loomInitialRow = 9;
                const int loomSheetRowLimit = 37;

                object[,] arrFinalReportLOOM = new object[maxRows, 17];
                Excel.Worksheet activeSheet = (Excel.Worksheet)modExcelInst.ExcelApp.ActiveWorkbook.ActiveSheet;

                // Bulk-read columns 1-7 for all needed rows in a single COM call instead of
                // 7 individual Cells[].Value2 reads per row. Range.Value2 for a multi-cell range
                // comes back 1-indexed on both dimensions, so bulkValues[row, col] lines up
                // directly with the original activeSheet.Cells[row + 1, col] access.
                object[,] bulkValues = null;
                if (maxRows >= 1)
                {
                    bulkValues = ExcelBulk.ReadValues(activeSheet, 2, 1, maxRows + 1, 7);
                }

                for (int row = 1; row <= maxRows; row++)
                {
                    string wireCode = Convert.ToString(bulkValues[row, 5]);
                    string sourcePin = Convert.ToString(bulkValues[row, 1]);
                    string sourceConn = Convert.ToString(bulkValues[row, 2]);
                    string destPin = Convert.ToString(bulkValues[row, 3]);
                    string destConn = Convert.ToString(bulkValues[row, 4]);
                    string destination = Convert.ToString(bulkValues[row, 7]);
                    string remarks = Convert.ToString(bulkValues[row, 6]);

                    if (row == 1 || !(wireCode == Convert.ToString(arrFinalReportLOOM[finalRowCount - 1, 2]) &&
                                      sourcePin == Convert.ToString(arrFinalReportLOOM[finalRowCount - 1, 4]) &&
                                      sourceConn == Convert.ToString(arrFinalReportLOOM[finalRowCount - 1, 5]) &&
                                      destPin == Convert.ToString(arrFinalReportLOOM[finalRowCount - 1, 7]) &&
                                      destConn == Convert.ToString(arrFinalReportLOOM[finalRowCount - 1, 8])))
                    {
                        arrFinalReportLOOM[finalRowCount, 2] = wireCode;
                        arrFinalReportLOOM[finalRowCount, 3] = destination;
                        arrFinalReportLOOM[finalRowCount, 4] = sourcePin;
                        arrFinalReportLOOM[finalRowCount, 5] = sourceConn;
                        arrFinalReportLOOM[finalRowCount, 6] = "";
                        arrFinalReportLOOM[finalRowCount, 7] = destPin;
                        arrFinalReportLOOM[finalRowCount, 8] = destConn;
                        arrFinalReportLOOM[finalRowCount, 9] = "";

                        string[] wireParts = wireCode?.Split('/');
                        if (wireParts?.Length >= 2)
                        {
                            arrFinalReportLOOM[finalRowCount, 10] = wireParts[1] + " " + remarks;
                        }

                        finalRowCount++;
                    }
                }

                int numberOfSheets = ExcelRowRules.GetNumberOfSheetsRequired(finalRowCount, loomSheetRowLimit);
                int currentSheet = 1;
                int serialNumber = 0;
                int startRow = 0;

                // A given row's wire code otherwise gets Split('/') at least twice - once as a
                // "compareRow" against the previous group's baseRow, and again as "baseRow" once
                // it becomes the next group's leader. Precompute each row's key once instead.
                var wireKeys = new string[finalRowCount];
                for (int r = 0; r < finalRowCount; r++)
                {
                    string[] parts = Convert.ToString(arrFinalReportLOOM[r, 2])?.Split('/');
                    wireKeys[r] = parts?.Length >= 2 ? parts[0] + parts[1] : "";
                }

                while (startRow < finalRowCount - 1)
                {
                    int baseRow = startRow;
                    if (baseRow > currentSheet * loomSheetRowLimit)
                        currentSheet++;

                    string baseWireKey = wireKeys[baseRow];

                    for (int compareRow = baseRow + 1; compareRow < finalRowCount; compareRow++)
                    {
                        string compareWireKey = wireKeys[compareRow];

                        arrFinalReportLOOM[baseRow, 13] = baseRow - (currentSheet - 1) * loomSheetRowLimit + loomInitialRow;

                        if (compareWireKey == baseWireKey)
                        {
                            arrFinalReportLOOM[baseRow, 12] = "ME";
                            arrFinalReportLOOM[baseRow, 14] = compareRow - (currentSheet - 1) * loomSheetRowLimit + loomInitialRow;
                            arrFinalReportLOOM[baseRow, 16] = arrFinalReportLOOM[baseRow, 10];
                            startRow = compareRow + 1;
                        }
                        else
                        {
                            if (!"ME".Equals(Convert.ToString(arrFinalReportLOOM[baseRow, 12])))
                            {
                                arrFinalReportLOOM[baseRow, 12] = "UN";
                                arrFinalReportLOOM[baseRow, 14] = baseRow - (currentSheet - 1) * loomSheetRowLimit + loomInitialRow;
                                arrFinalReportLOOM[baseRow, 16] = arrFinalReportLOOM[baseRow, 10];
                                startRow = compareRow;
                            }
                            break;
                        }
                    }

                    serialNumber++;
                    arrFinalReportLOOM[baseRow, 15] = serialNumber;
                }

                if (arrFinalReportLOOM[finalRowCount - 1, 12] == null)
                {
                    arrFinalReportLOOM[finalRowCount - 1, 12] = "LAST";
                    arrFinalReportLOOM[finalRowCount - 1, 15] = serialNumber + 1;
                }

                modExcelInst.ExcelApp.ActiveWorkbook.Close(false);
                modExcelInst.GenerateHALReportFormat_CableList(arrFinalReportLOOM, reportName, numberOfSheets, loomSheetRowLimit, ConfigurationManager.AppSettings["PanelDrawingFolder"]);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error in CableList_Report_Preprocessing: " + ex.Message);
            }
        }

        /* public static Dictionary<string, string> connectorMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
         {
             { "J1", "A" }, { "A", "J1" },
             { "J2", "B" }, { "B", "J2" },
             { "J3", "C" }, { "C", "J3" },
             { "J4", "D" }, { "D", "J4" },
             { "J5", "E" }, { "E", "J5" },
             { "J6", "F" }, { "F", "J6" },
             { "J7", "G" }, { "G", "J7" },
             { "J8", "H" }, { "H", "J8" },
             { "J9", "J" }, { "J", "J9" },   // Skipped 'I'
             { "J10", "K" }, { "K", "J10" },
             { "J11", "L" }, { "L", "J11" },
             { "J12", "M" }, { "M", "J12" },
             { "J13", "N" }, { "N", "J13" },
             { "J14", "P" }, { "P", "J14" }, // Skipped 'O'
             { "J15", "Q" }, { "Q", "J15" },
             { "J16", "R" }, { "R", "J16" },
             { "J17", "S" }, { "S", "J17" },
             { "J18", "T" }, { "T", "J18" },
             { "J19", "U" }, { "U", "J19" },
             { "J20", "V" }, { "V", "J20" },
             { "J21", "W" }, { "W", "J21" },
             { "J22", "X" }, { "X", "J22" },
             { "J23", "Y" }, { "Y", "J23" },
             { "J24", "Z" }, { "Z", "J24" },
         };*/

        /*  public void CableList_Report_Preprocessing(object[,] iarr, string reportName, modExcel modExcelInst)
          {
              try
              {
                  int MaxARR = ExcelRowRules.MaxRowinArray(iarr);

                  int Z = 1;
                  object[,] arrCustomLoom = new object[MaxARR, 9];
                  object[,] arrFinalReportLOOM = new object[MaxARR, 16];
                  string[] arr22;
                  int O, P, c = 1;
                  Excel.Worksheet activeSheet = (Excel.Worksheet)modExcelInst.ExcelApp.ActiveWorkbook.ActiveSheet;

                  for (O = 1; O <= MaxARR; O++)
                  {
                      // Avoid duplication of the same connection and wirecode
                      if (O == 1)
                      {
                          arrFinalReportLOOM[c, 2] = activeSheet.Cells[O + 1, 5].Value;
                          arrFinalReportLOOM[c, 3] = activeSheet.Cells[O + 1, 7].Value;
                          arrFinalReportLOOM[c, 4] = activeSheet.Cells[O + 1, 1].Value;
                          arrFinalReportLOOM[c, 5] = activeSheet.Cells[O + 1, 2].Value;
                          arrFinalReportLOOM[c, 6] = "";  // Empty cell for placeholder
                          arrFinalReportLOOM[c, 7] = activeSheet.Cells[O + 1, 3].Value;
                          arrFinalReportLOOM[c, 8] = activeSheet.Cells[O + 1, 4].Value;
                          arrFinalReportLOOM[c, 9] = "";  // Empty cell for placeholder

                          // Splitting string by '/' and assigning to array
                          arr22 = arrFinalReportLOOM[c, 2].ToString().Split('/');
                          arrFinalReportLOOM[c, 10] = arr22[1] + " " + activeSheet.Cells[O + 1, 6].Value.ToString();
                          c++;
                      }
                      else if (activeSheet.Cells[O + 1, 5].Value == arrFinalReportLOOM[c - 1, 2].ToString())
                      {
                          if (activeSheet.Cells[O + 1, 1].Value == arrFinalReportLOOM[c - 1, 7].ToString() &&
                              activeSheet.Cells[O + 1, 2].Value == arrFinalReportLOOM[c - 1, 8].ToString() &&
                              activeSheet.Cells[O + 1, 3].Value == arrFinalReportLOOM[c - 1, 4].ToString() &&
                              activeSheet.Cells[O + 1, 4].Value == arrFinalReportLOOM[c - 1, 5].ToString())
                          {
                              // Do nothing if the condition matches
                          }
                          else
                          {
                              arrFinalReportLOOM[c, 2] = activeSheet.Cells[O + 1, 5].Value;
                              arrFinalReportLOOM[c, 3] = activeSheet.Cells[O + 1, 7].Value;
                              arrFinalReportLOOM[c, 4] = activeSheet.Cells[O + 1, 1].Value;
                              arrFinalReportLOOM[c, 5] = activeSheet.Cells[O + 1, 2].Value;
                              arrFinalReportLOOM[c, 6] = "";  // Empty cell for placeholder
                              arrFinalReportLOOM[c, 7] = activeSheet.Cells[O + 1, 3].Value;
                              arrFinalReportLOOM[c, 8] = activeSheet.Cells[O + 1, 4].Value;
                              arrFinalReportLOOM[c, 9] = "";  // Empty cell for placeholder

                              arr22 = arrFinalReportLOOM[c, 2].ToString().Split('/');
                              arrFinalReportLOOM[c, 10] = arr22[1] + " " + activeSheet.Cells[O + 1, 6].Value.ToString();
                              c++;
                          }
                      }
                      else
                      {
                          arrFinalReportLOOM[c, 2] = activeSheet.Cells[O + 1, 5].Value;
                          arrFinalReportLOOM[c, 3] = activeSheet.Cells[O + 1, 7].Value;
                          arrFinalReportLOOM[c, 4] = activeSheet.Cells[O + 1, 1].Value;
                          arrFinalReportLOOM[c, 5] = activeSheet.Cells[O + 1, 2].Value;
                          arrFinalReportLOOM[c, 6] = "";  // Empty cell for placeholder
                          arrFinalReportLOOM[c, 7] = activeSheet.Cells[O + 1, 3].Value;
                          arrFinalReportLOOM[c, 8] = activeSheet.Cells[O + 1, 4].Value;
                          arrFinalReportLOOM[c, 9] = "";  // Empty cell for placeholder

                          arr22 = arrFinalReportLOOM[c, 2].ToString().Split('/');
                          arrFinalReportLOOM[c, 10] = arr22[1] + " " + activeSheet.Cells[O + 1, 6].Value.ToString();
                          c++;
                      }
                  }

                  // Define additional variables for sheet creation
                  int LoomSheetRowRequired = 37;
                  int NoLoomSheets = ExcelRowRules.GetNumberOfSheetsRequired(c, LoomSheetRowRequired);

                  int SlNo = 0;
                  int StartupRow = 1;
                  int SheetNum = 1;

                  while (StartupRow < c - 1)
                  {
                      int S1 = StartupRow;
                      if (S1 > SheetNum * LoomSheetRowRequired)
                          SheetNum++;

                      arr22 = arrFinalReportLOOM[S1, 2].ToString().Split('/');
                      string sWireCode1 = arr22[0] + arr22[1];
                      string sWireCode2;

                      for (int T = S1 + 1; T < c - 1; T++)
                      {
                          arr22 = arrFinalReportLOOM[T, 2].ToString().Split('/');
                          sWireCode2 = arr22[0] + arr22[1];

                          arrFinalReportLOOM[S1, 13] = S1 - (SheetNum - 1) * LoomSheetRowRequired + 9;

                          if (sWireCode2 == sWireCode1)
                          {
                              arrFinalReportLOOM[S1, 12] = "ME";
                              arrFinalReportLOOM[S1, 14] = T - (SheetNum - 1) * LoomSheetRowRequired + 9;
                              arrFinalReportLOOM[S1, 16] = arrFinalReportLOOM[S1, 10];
                              StartupRow = T + 1;
                          }
                          else
                          {
                              if (arrFinalReportLOOM[S1, 12].ToString() != "ME")
                              {
                                  arrFinalReportLOOM[S1, 12] = "UN";
                                  arrFinalReportLOOM[S1, 14] = S1 - (SheetNum - 1) * LoomSheetRowRequired + 9;
                                  arrFinalReportLOOM[S1, 16] = arrFinalReportLOOM[S1, 10];
                                  StartupRow = T;
                              }
                              break;
                          }
                      }

                      SlNo++;
                      arrFinalReportLOOM[S1, 15] = SlNo;
                  }

                  // If the last row is empty, mark it as "LAST"
                  if (arrFinalReportLOOM[c - 1, 12] == null || arrFinalReportLOOM[c - 1, 12].ToString() == "")
                  {
                      arrFinalReportLOOM[c - 1, 12] = "LAST";
                      arrFinalReportLOOM[c - 1, 15] = SlNo + 1;
                  }

                  // Close the Excel workbook
                  modExcelInst.ExcelApp.ActiveWorkbook.Close();

                  // Call the report generation function
                  modExcelInst.GenerateHALReportFormat_CableList(arrFinalReportLOOM, reportName, NoLoomSheets, LoomSheetRowRequired, ConfigurationManager.AppSettings["BundleFolder"]);
              }
              catch (Exception ex)
              {
                  Console.WriteLine("An error occurred: " + ex.Message);
              }
          }*/

    }

}
