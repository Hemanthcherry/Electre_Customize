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
using System.Text.RegularExpressions;
using System.Windows.Forms;
using static System.Runtime.InteropServices.JavaScript.JSType;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;
using Application = System.Windows.Forms.Application;
using Excel = Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;

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
                for (i = 0; i < ElecCollection.Count; i++)
                {
                    ElectreObj = ElecCollection[i];

                    SearchAndAppend(ElectreObj.ConnectorName, ref arrListOfComponents);
                    SearchAndAppend(ElectreObj.BundleName, ref arrListOfLOOM);
                    SearchAndAppend(ElectreObj.SheetName, ref arrListOfSHEET);
                    SearchAndAppend(ElectreObj.Panel, ref arrListOfPANEL);

                    switch (ElectreObj.ComponentType)
                    {
                        case "TBK":
                            SearchAndAppend(ElectreObj.ConnectorName, ref arrListOfJUN);
                            SearchAndAppend(ElectreObj.ConnectorName, ref arrListOfJUNnSPL); // For JM Link
                            break;
                        case "EQU":
                            SearchAndAppend(ElectreObj.ConnectorName, ref arrListOfEQU); // For Connector
                            break;
                        case "REL":
                            SearchAndAppend(ElectreObj.ConnectorName, ref arrListOfREL);   // For Realy
                            break;
                        case "DIS":
                            SearchAndAppend(ElectreObj.ConnectorName, ref arrListOfDIS);  // For Break Connector
                            break;
                        case "SPL":
                            SearchAndAppend(ElectreObj.ConnectorName, ref arrListOfSPL);
                            SearchAndAppend(ElectreObj.ConnectorName, ref arrListOfJUNnSPL); // For JM Link
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
                            SearchAndAppend(ElectreObj.ConnectorName, ref arrListOfSWT);
                            break;
                    }

                    // List the number of pins in each connector
                    if (ElectreObj.ComponentType == "EQU")
                    {
                        SearchAndAppend($"{ElectreObj.ConnectorName};{ElectreObj.PinNumber}", ref arrConnectorNAMEandPIN);
                    }
                }

                arrListOfComponentsInPANEL = new object[arrListOfComponents.Count, 4];

                // sorting lists by order
                arrListOfLOOM = arrListOfLOOM.OrderBy(x => x).ToList();
                arrListOfSHEET = arrListOfSHEET.OrderBy(x => x).ToList();
                arrListOfComponents = arrListOfComponents.OrderBy(x => x).ToList();
                arrListOfJUN = arrListOfJUN.OrderBy(x => x).ToList();
                arrListOfJUNnSPL = arrListOfJUNnSPL.OrderBy(x => x).ToList();
                arrListOfSPL = arrListOfSPL.OrderBy(x => x).ToList();
                arrListOfEQU = arrListOfEQU.OrderBy(x => x).ToList();
                arrListOfREL = arrListOfREL.OrderBy(x => x).ToList();
                arrListOfDIS = arrListOfDIS.OrderBy(x => x).ToList();
                arrListOfSWT = arrListOfSWT.OrderBy(x => x).ToList();

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
                for (i = 0; i < ElecCollection_Panel.Count; i++)
                {
                    ElectreObj = ElecCollection_Panel[i];

                    // SearchAndAppend(ElectreObj.ConnectorName, ref arrListOfComponents);
                    // SearchAndAppend(ElectreObj.BundleName, ref arrListOfLOOM);
                    // SearchAndAppend(ElectreObj.SheetName, ref arrListOfSHEET);
                    SearchAndAppend(ElectreObj.Panel, ref arrListOfPANEL);

                    switch (ElectreObj.ComponentType)
                    {
                        case "TBK":
                            //  SearchAndAppend(ElectreObj.ConnectorName, ref arrListOfJUN);
                            //  SearchAndAppend(ElectreObj.ConnectorName, ref arrListOfJUNnSPL); // For JM Link
                            break;
                        case "EQU":
                            // SearchAndAppend(ElectreObj.ConnectorName, ref arrListOfEQU);
                            break;
                        case "REL":
                            // SearchAndAppend(ElectreObj.ConnectorName, ref arrListOfREL);
                            break;
                        case "DIS":
                            // SearchAndAppend(ElectreObj.ConnectorName, ref arrListOfDIS);
                            break;
                        case "SPL":
                            // SearchAndAppend(ElectreObj.ConnectorName, ref arrListOfSPL);
                            // SearchAndAppend(ElectreObj.ConnectorName, ref arrListOfJUNnSPL); // For JM Link
                            break;
                        case "SWT":
                        case "ERM":
                        case "IND":
                        case "SCB":
                            SearchAndAppend(ElectreObj.ConnectorName, ref arrPowerOnCircuitBreakers);
                            SearchAndAppend(ElectreObj.ComponentType, ref arrPowerOnCircuitBreakersType);
                            break;
                        case "TCB":
                            SearchAndAppend(ElectreObj.ConnectorName, ref arrPowerOnCircuitBreakers);
                            SearchAndAppend(ElectreObj.ComponentType, ref arrPowerOnCircuitBreakersType);
                            break;
                        case "TER":
                        case "FUS":
                        case "POT":
                        case "LMP":
                        case "ANT":
                        case "BUS":
                        case "MSW":
                            //SearchAndAppend(ElectreObj.ConnectorName, ref arrListOfSWT);
                            break;
                    }

                    // List the number of pins in each connector
                    if (ElectreObj.ComponentType == "EQU")
                    {
                        SearchAndAppend($"{ElectreObj.ConnectorName};{ElectreObj.PinNumber}", ref arrConnectorNAMEandPIN);
                    }
                }               

                ElecCollection_All = ElecCollection_All.OrderBy(E => E.ConnectorName).ThenBy(e => e.PinNumber).ToList();

                arrListOfComponentsInPANEL = new object[arrListOfComponents.Count, 4];

                arrListOfPANEL = arrListOfPANEL.OrderBy(x => x).ToList();

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
        // This method adds components to a List object
        public static bool SearchAndAppend(string istrName, ref List<string> iarr)
        {
            for (int J = 0; J < iarr.Count; J++)
            {
                if (iarr[J].Equals(istrName))
                {
                    // Item found, nothing to add
                    return false; // Returns if anything was added or not
                }
            }

            // Item not found, resize the array and add the new item           
            iarr.Add(istrName);

            return true; // Returns if anything was added or not
        }

        //this method assigning data from Elecollection object to 2d array arrFT_CwithBC and this arrFT_CwithBC Used for Componentes specifice breakdown reports generation purpose
        public static void Removing_DuplicateWires_In2DArray()
        {
            try
            {
                CwithBCReportRow = 0;
                Logging.Info($"Removing_DuplicateWires_In2DArray started");
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

                        string coreNumber = IsValidCoreNumber(E1.Core_Part_Number) ? E1.Core_Part_Number : "";
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

                        int j = 0;
                        do
                        {
                            if (j != i) // To omit this line
                            {
                                // Removed signal == condition HAL
                                if (ElecCollection[j].WireNumber.ToUpper() == E1.WireNumber.ToUpper() &&
                                     ElecCollection[j].SubNet.ToUpper() == E1.SubNet.ToUpper() &&
                                     ElecCollection[j].ComponentType != "SDS" &&
                                        !string.IsNullOrEmpty(ElecCollection[j].WireNumber))
                                {
                                    arrFT_CwithBC[CwithBCReportRow, TC1] = ElecCollection[j].ConnectorName;
                                    arrFT_CwithBC[CwithBCReportRow, TP1] = ElecCollection[j].PinNumber;
                                    arrFT_CwithBC[CwithBCReportRow, 10] = layerAssign(ElecCollection[j].Layer);

                                    // Logging.Info($"{CwithBCReportRow}: {arrFT_CwithBC[CwithBCReportRow, 0]}--{arrFT_CwithBC[CwithBCReportRow, 1]}--{arrFT_CwithBC[CwithBCReportRow, 2]}--{arrFT_CwithBC[CwithBCReportRow, 3]}");

                                    CwithBCReportRow++;
                                }
                            }
                            j++;
                        } while (j < ElecCollection.Count);
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

                        string coreNumber = IsValidCoreNumber(E1.Core_Part_Number) ? E1.Core_Part_Number : "";

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

                        int j = 0;
                        do
                        {
                            if (i != j)
                            {
                                if (ElecCollection[j].WireNumber.ToUpper() == E1.WireNumber.ToUpper() &&
                                     ElecCollection[j].SubNet.ToUpper() == E1.SubNet.ToUpper() &&
                                     ElecCollection[j].ComponentType != "SDS" &&
                                        !string.IsNullOrEmpty(ElecCollection[j].WireNumber) &&
                                        ElecCollection[j].Tag6_Link != "Linked")
                                {
                                    arrFT_CwithBCProject[CwithBCReportRowProject, TC1] = ElecCollection[j].ConnectorName;
                                    arrFT_CwithBCProject[CwithBCReportRowProject, TP1] = ElecCollection[j].PinNumber;
                                    arrFT_CwithBCProject[CwithBCReportRowProject, 9] = ElecCollection[j].SheetName;
                                    arrFT_CwithBCProject[CwithBCReportRowProject, 10] = ElecCollection[j].BundleName;
                                    ElecCollection[j].Tag6_Link = "Linked";
                                    arrFT_CwithBCProject[CwithBCReportRowProject, 12] = layerAssign(ElecCollection[j].Layer);
                                    arrFT_CwithBCProject[CwithBCReportRowProject, 13] = ElecCollection[j].Group;

                                    // Logging.Info($"{arrFT_CwithBCProject[CwithBCReportRowProject, 0]}--{arrFT_CwithBCProject[CwithBCReportRowProject, 1]}--{arrFT_CwithBCProject[CwithBCReportRowProject, 2]}--{arrFT_CwithBCProject[CwithBCReportRowProject, 3]}");
                                    CwithBCReportRowProject++;
                                }
                            }
                            j++;
                        } while (j < ElecCollection.Count);

                    }
                }

                Logging.Info("Reading_And_StoringData_In2DArray completed");
            }

            catch (Exception ex)
            {
                MessageBox.Show("modMain #002: " + ex.Message);
            }
        }
        private static bool IsValidCoreNumber(string tag3)
        {
            return int.TryParse(tag3, out int coreNumber) && coreNumber >= 1 && coreNumber <= 18;
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
                    }

                    if (_customReportWindow.SelectedLoomList.Count > 0)
                    {
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
                    }

                    if (_customReportWindow.SelectedSheetList.Count > 0)
                    {
                        string fullPath = Path.Combine(templpath, wireListSheetFolder);
                       // DeleteExistingFiles(fullPath);

                        sheetsReportWireList(modExcelInst, _customReportWindow.SelectedSheetList, wireListSheetFolder);

                        // Update loading message on the UI thread
                        _loadingForm.Invoke(() =>
                        {
                            _loadingForm.LbsLoadinMessag = "Sheet Report Generated";
                        });
                    }

                    // Component Breakdown
                    if (_customReportWindow.SelectedEqupList.Count > 0)
                    {
                        string fullPath = Path.Combine(templpath, EquipmentFolder);
                        //DeleteExistingFiles(fullPath);

                        equipmentReportWireList(modExcelInst, "COMP_WireList", ConfigurationManager.AppSettings["EquipmentFolder"], _customReportWindow.SelectedEqupList);

                        // Update loading message on the UI thread
                        _loadingForm.Invoke(() =>
                        {
                            _loadingForm.LbsLoadinMessag = "Equipment Report Generated";
                        });
                    }

                    if (_customReportWindow.SelectedBrkList.Count > 0)
                    {
                        string fullPath = Path.Combine(templpath, BrkFolder);
                        //DeleteExistingFiles(fullPath);

                        equipmentReportWireList(modExcelInst, "BRK_WireList", ConfigurationManager.AppSettings["BrkFolder"], _customReportWindow.SelectedBrkList);

                        // Update loading message on the UI thread
                        _loadingForm.Invoke(() =>
                        {
                            _loadingForm.LbsLoadinMessag = "Break Down Report Generated";
                        });
                    }

                    if (_customReportWindow.SelectedJmList.Count > 0)
                    {
                        string fullPath = Path.Combine(templpath, JmFolder);
                        //DeleteExistingFiles(fullPath);

                        equipmentReportWireList(modExcelInst, "JM_WireList", ConfigurationManager.AppSettings["JmFolder"], _customReportWindow.SelectedJmList);

                        // Update loading message on the UI thread
                        _loadingForm.Invoke(() =>
                        {
                            _loadingForm.LbsLoadinMessag = "Junction Module Report Generated";
                        });
                    }

                    if (_customReportWindow.SelectedMiscList.Count > 0)
                    {
                        string fullPath = Path.Combine(templpath, MiscFolder);
                        //DeleteExistingFiles(fullPath);

                        equipmentReportWireList(modExcelInst, "MISC_WireList", ConfigurationManager.AppSettings["MiscFolder"], _customReportWindow.SelectedMiscList);

                        // Update loading message on the UI thread
                        _loadingForm.Invoke(() =>
                        {
                            _loadingForm.LbsLoadinMessag = "MISC Report Generated";
                        });
                    }

                    //Continuity
                    if (_customReportWindow.SelectedContLoomList.Count > 0)
                    {
                        ContinuityWOBreakdown cwob = new ContinuityWOBreakdown(modExcelInst, _customReportWindow.SelectedContLoomList);
                        cwob.ContinuityReportGeneration(ContinuityFolder, ContinuityCompFolder, ElecCollection);
                        // Update loading message on the UI thread
                        _loadingForm.Invoke(() =>
                        {
                            _loadingForm.LbsLoadinMessag = "Continuity Report Generated";
                        });
                    }

                    // Megger
                    if (_customReportWindow.SelectedMeggerLoomList.Count > 0)
                    {
                        MeggerSchedulesheet(modExcelInst, _customReportWindow.SelectedMeggerLoomList, MeggerFolder, ElecCollection);

                        // Update loading message on the UI thread
                        _loadingForm.Invoke(() =>
                        {
                            _loadingForm.LbsLoadinMessag = "Megger Report Generated";
                        });
                    }

                    // Power On
                    if (_customReportWindow.powerOnProjectList)
                    {
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
                    }

                    if (_customReportWindow.SelectedPanelList.Count > 0)
                    {
                        string fullPath = Path.Combine(templpath, PowerOnFolder);
                        //DeleteExistingFiles(fullPath);

                        var SelectedPanels = _customReportWindow.SelectedPanelList;
                        foreach (var panel in SelectedPanels)
                        {
                            string sanitizedSheetName = SanitizeSheetName(panel);
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
                    }

                    if (_customReportWindow.SelectedSheetListPanelDwgCL.Count > 0)
                    {
                        string fullPath = Path.Combine(templpath, PanelSheetFolderCL);
                        //DeleteExistingFiles(fullPath);

                        sheetsReportWireList(modExcelInst, _customReportWindow.SelectedSheetListPanelDwgCL, PanelSheetFolderCL);

                        // Update loading message on the UI thread
                        _loadingForm.Invoke(() =>
                        {
                            _loadingForm.LbsLoadinMessag = "Sheet Report Generated";
                        });
                    }

                    // Panel Schedules Continuity
                    if (_customReportWindow.SelectedLoomListPanelDwgCont.Count > 0)
                    {
                        ContinuityWOBreakdown cwob = new ContinuityWOBreakdown(modExcelInst, _customReportWindow.SelectedLoomListPanelDwgCont);
                        cwob.ContinuityReportGeneration(PanelContinuityFolder, PanelContinuityCompFolder, PanelDrawingSchedulesWindow.filteredSheetCollection);
                        // Update loading message on the UI thread
                        _loadingForm.Invoke(() =>
                        {
                            _loadingForm.LbsLoadinMessag = "Continuity Report Generated";
                        });
                    }

                    // Panel Schedules Megger
                    if (_customReportWindow.SelectedLoomListPanelDwgMeg.Count > 0)
                    {
                        MeggerSchedulesheet(modExcelInst, _customReportWindow.SelectedLoomListPanelDwgMeg, PanelMeggerFolder, PanelDrawingSchedulesWindow.filteredSheetCollection);

                        // Update loading message on the UI thread
                        _loadingForm.Invoke(() =>
                        {
                            _loadingForm.LbsLoadinMessag = "Megger Report Generated";
                        });
                    }

                    // Panel Schedules Power On
                    if (_customReportWindow.SelectedSheetListPanelDwgPower.Count > 0)
                    {
                        string fullPath = Path.Combine(templpath, PanelPowerOnFolder);
                        if (!Directory.Exists(fullPath))
                        {
                            Directory.CreateDirectory(fullPath);
                        }
                       // DeleteExistingFiles(fullPath);

                        var SelectedSheets = _customReportWindow.SelectedSheetListPanelDwgPower;
                        foreach (var sheet in SelectedSheets)
                        {
                            string sanitizedSheetName = SanitizeSheetName(sheet);
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

        public static void DeleteExistingFiles(string directoryPath)
        {
            if (!Directory.Exists(directoryPath))
            {
                Logging.Warning("Directory does not exist: " + directoryPath);
                return;
            }

            string[] extensionsToDelete = { ".xls", ".xlsx", ".txt", ".csv" };

            try
            {
                var files = Directory.GetFiles(directoryPath)
                                     .Where(file => extensionsToDelete.Contains(Path.GetExtension(file), StringComparer.OrdinalIgnoreCase));

                foreach (var file in files)
                {
                    try
                    {
                        File.Delete(file);
                        Logging.Info("Deleted: " + file);
                    }
                    catch (Exception ex)
                    {
                        Logging.Error($"Error deleting file {file}: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Logging.Error($"Failed to retrieve files: {ex.Message}");
            }
        }

        // This method generates Excel reports for sheets in Cable List.
        private void sheetsReportWireList(modExcel modExcelInst, List<string> selectedSheetList, string sheetFolderPath)
        {
            try
            {
                //ElectreObject E17;
                object[,] arrCustomLoom = new object[arrFT_CwithBC.GetUpperBound(0), 10];
                Logging.Info("sheet wise Report generation Started");

                int z = 0;

                //foreach (string s in _mainForm.ListSHEET)
                foreach (string s in selectedSheetList)
                {
                    try
                    {
                        Logging.Info($"{s} Sheet report Started");
                       // bool appendPage = false;

                        for (int y = 0; y < arrFT_CwithBCProject.GetLength(0); y++)
                        {
                            if (arrFT_CwithBCProject[y, 9] != null)
                            {
                                if (arrFT_CwithBCProject[y, 9].ToString() == s)
                                {
                                    // Copy elements from arrFT_CwithBCProject to arrCustomLoom
                                    for (int x = 0; x < 9; x++)
                                    {
                                        arrCustomLoom[z, x] = arrFT_CwithBCProject[y, x];
                                    }
                                    arrCustomLoom[z, 9] = arrFT_CwithBCProject[y, 12];
                                    z++;
                                }
                            }
                        }

                        string sheetName = SanitizeSheetName(s);
                        // string workbookPath = modExcelInst.CreateNewWorkbook($"{s}_SHEET_Wirelist", sheetName, ConfigurationManager.AppSettings["SheetFolder"]);
                        string workbookPath = modExcelInst.CreateNewWorkbook($"{s}_SHEET_Wirelist", sheetName, sheetFolderPath);

                        if (!modExcelInst.CreateWirelistReportHeader(workbookPath))
                        {
                            Logging.Error($"Error Sheet report Header {s}");
                        }

                        modExcelInst.AppendToExcel(arrCustomLoom);
                        modExcelInst.SortWirelistWireNumber(2, arrCustomLoom.GetUpperBound(0) + 1);
                        z = 0;
                        arrCustomLoom = new object[arrFT_CwithBC.GetUpperBound(0), 10];
                        Logging.Info($"{s} Sheet report Completed");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"{s}:{ex.Message}");
                        Logging.Error($"{s}:{ex.Message}");
                    }
                }
            }

            catch (Exception ex)
            {
                MessageBox.Show("modMain #003: " + ex.Message);
            }
        }

        private string SanitizeSheetName(string sheetName)
        {
            // Remove invalid characters
            string invalidChars = new string(Path.GetInvalidFileNameChars()) + @":\/?*[]";
            foreach (char c in invalidChars)
            {
                sheetName = sheetName.Replace(c.ToString(), "");
            }

            // Trim to 31 characters
            return sheetName.Length > 31 ? sheetName.Substring(0, 31) : sheetName;
        }

        //this method gennerates Excel reports for project WireList
        private void WireListReport(modExcel modExcelInst)
        {
            object[,] arrCustomLoom = new object[arrFT_CwithBCProject.Length, 10];
            int Z = 0;

            for (int Y = 0; Y < arrFT_CwithBCProject.GetLength(0); Y++)
            {
                // Check if the value at position (Y, 9) matches the selected item in frmMain.listLOOM at index S
                for (int X = 0; X < 9; X++)  // Copy elements from arrFT_CwithBCProject to arrCustomLoom
                {
                    arrCustomLoom[Z, X] = arrFT_CwithBCProject[Y, X];
                }

                arrCustomLoom[Z, 9] = arrFT_CwithBCProject[Y, 12];

                Z++;
            }
            modExcelInst.AppendToExcel(arrCustomLoom);
            modExcelInst.SortWirelistWireNumber(2, arrFT_CwithBCProject.GetLength(0) + 1);

        }

        //this method generates Excel reports for Component Specific Breakdown 
        private void equipmentReportWireList(modExcel modExcelInst, string fileSuffixName, string FolderName, List<string> componentList)
        {
            try
            {
                ElectreObject E17;

                object[,] arrCustomLoom = new object[arrFT_CwithBCProject.Length, 10];

                int z = 0;

                //foreach (string s in _mainForm.ListEQU2)
                foreach (string s in componentList)
                {
                    try
                    {
                        Logging.Info($"{s} Equipment report Started");
                        for (int y = 0; y < arrFT_CwithBC.GetLength(0); y++)
                        {
                            if (arrFT_CwithBC[y, 0] != null)
                            {
                                if (arrFT_CwithBC[y, 0].ToString() == s)
                                {
                                    // Copy elements from arrFT_CwithBC to arrCustomLoom
                                    for (int x = 0; x < 9; x++)
                                    {
                                        arrCustomLoom[z, x] = arrFT_CwithBC[y, x];
                                    }
                                    arrCustomLoom[z, 9] = arrFT_CwithBC[y, 10];
                                    z++;
                                }
                            }
                        }

                        string workbookPath = modExcelInst.CreateNewWorkbook($"{s}_{fileSuffixName}", s, FolderName);

                        if (!modExcelInst.CreateWirelistReportHeader(workbookPath))
                        {
                            Logging.Error($"Error Report Header {s}");
                        }

                        modExcelInst.AppendToExcel(arrCustomLoom);
                        modExcelInst.SortWirelistWireNumber(2, arrCustomLoom.GetUpperBound(0) + 1);
                        z = 0;
                        arrCustomLoom = new object[arrFT_CwithBC.GetUpperBound(0), 10];
                        Logging.Info($"{s} Equipment report Completed");
                    }
                    catch (Exception ex)
                    {

                        MessageBox.Show($"{s}:{ex.Message}");
                        Logging.Error($"{s}:{ex.Message}");

                    }

                }
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
                //foreach (string loomName in _mainForm.ListLOOM)
                foreach (string loomName in selectedLoomList)
                {
                    try
                    {
                        Logging.Info($"{loomName} Loom report without template Started");
                        object[,] arrCustomLoom = new object[arrFT_CwithBCProject.Length, 14];
                        int Z = 0;

                        for (int Y = 0; Y < arrFT_CwithBCProject.GetLength(0); Y++)
                        {
                            // ✅ New filter: if sheet filtering is active, skip rows not in selected sheets
                            string sheetName = arrFT_CwithBCProject[Y, 7]?.ToString(); // adjust index if sheet column is elsewhere
                            if (selectedSheets != null && !selectedSheets.Contains(sheetName))
                            {
                                continue;
                            }
                            // Check if the value at position (Y, 9) matches the selected item in frmMain.listLOOM at index S
                            if (arrFT_CwithBCProject[Y, 8]?.ToString() == loomName)  // Arrays are zero-indexed in C#
                            {
                                for (int X = 0; X < 9; X++)  // Copy elements from arrFT_CwithBCProject to arrCustomLoom
                                {
                                    arrCustomLoom[Z, X] = arrFT_CwithBCProject[Y, X];
                                }

                                arrCustomLoom[Z, 9] = arrFT_CwithBCProject[Y, 12];
                                arrCustomLoom[Z, 10] = arrFT_CwithBCProject[Y, 13];

                                Z++;
                            }
                            else if (arrFT_CwithBCProject[Y, 10]?.ToString() == loomName)
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
                        loomSortList.Distinct().ToList();

                        // LINQ query to exclude groups that not start with an alphabet
                        var groupnumbers = loomSortList
                                        .Where(x => !string.IsNullOrEmpty(x.Group) && Regex.IsMatch(x.Group, @"^[A-Za-z]"))
                                        .Select(x => new { x.Group, x.WireCode })
                                        .ToList();

                        // ✅ Regex pattern to match valid cable types (numbers with _, - or just digits)
                        Regex validPattern = new Regex(@"^[\dA-Za-z_-]+$");

                        var otherCableTypes = loomSortList
                                        .Where(x => !string.IsNullOrEmpty(x.Group) && !validPattern.IsMatch(x.Group))
                                        .Select(x => new { x.Group, x.WireCode })
                                        .ToList();

                        // Combine both lists
                        var combinedList = groupnumbers.Concat(otherCableTypes).ToList();

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

                        // if group number start with a letters
                        var filteredLoomSortList = loomSortList
                      .Where(x => !Regex.IsMatch(x.Group, @"^[A-Za-z]")) // Exclude groups starting with a letter
                        .ToList();

                        loomSortList = filteredLoomSortList
                      .OrderBy(x =>
                      {
                          // Extract main number from the beginning of the string
                          var match = Regex.Match(x.Group, @"^(\d+)");
                          return match.Success ? int.Parse(match.Groups[1].Value) : int.MaxValue;
                      })
                        .ThenBy(x =>
                        {
                            // Extract alphabetical part after the number (case-insensitive)
                            var match = Regex.Match(x.Group, @"^(\d+)([A-Za-z]*)");
                            return match.Success ? match.Groups[2].Value.ToLower() : "";
                        })
                     .ThenBy(x =>
                     {
                         // Extract sub-number after _ or - (default to 0 if not found)
                         var parts = Regex.Split(x.Group.Replace('_', '-'), "[-_]");
                         return parts.Length > 1 && int.TryParse(parts[1], out int subNumber) ? subNumber : 0;
                     })
                        .ThenBy(x => x.WireCode)  // Sort by WireCode last
                      .ToList();

                        int loomRowCount = 0;

                        for (int i = 0; i < arrCustomLoom.GetLength(0); i++)
                        {
                            if (arrCustomLoom[i, 0] != null && arrCustomLoom[i, 4] != null)
                            {
                                loomRowCount++;
                            }
                        }
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
                        string workbookPath = modExcelInst.CreateNewWorkbook($"{loomName}_Loom_Wirelist", loomName, loomFolderPath);

                        if (!modExcelInst.CreateloomlistReportHeader(workbookPath))
                        {
                            Logging.Error($"Error Sheet report Header {loomName}");
                        }

                        modExcelInst.AppendToExcel(arrFinalReportLOOM);
                        // modExcelInst.SortWirelistWireNumber3(2, arrFinalReportLOOM.GetUpperBound(0) + 1);
                        Logging.Info($"{loomName} Loom report without Template Completed");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"{loomName}:incorrect cable group id");
                        Logging.Error($"{loomName}: {ex.Message}");
                    }
                }
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
                List<string> allspecialcables = new List<string>();
                allspecialcables = ElecCollection.Where(p => p.CableType == "X").Select(p => p.Tag7).ToList();

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


                //foreach (string loomName in _mainForm.ListLOOM)
                foreach (string loomName in selectedLoomList)
                {
                    try
                    {
                        Logging.Info($"{loomName} Loom report Started");
                        object[,] arrCustomLoom = new object[arrFT_CwithBCProject.Length, 14];
                        int Z = 0;

                        for (int Y = 0; Y < arrFT_CwithBCProject.GetLength(0); Y++)
                        {
                            // ✅ New filter: check sheet name if required
                            string sheetName = arrFT_CwithBCProject[Y, 7]?.ToString(); // adjust index if needed
                            if (selectedSheets != null && !selectedSheets.Contains(sheetName))
                                continue;

                            // Check if the value at position (Y, 9) matches the selected item in frmMain.listLOOM at index S
                            if (arrFT_CwithBCProject[Y, 8]?.ToString() == loomName)  // Arrays are zero-indexed in C#
                            {
                                for (int X = 0; X < 13; X++)  // Copy elements from arrFT_CwithBCProject to arrCustomLoom
                                {
                                    arrCustomLoom[Z, X] = arrFT_CwithBCProject[Y, X];
                                }
                                arrCustomLoom[Z, 13] = arrFT_CwithBCProject[Y, 13];

                                Z++;
                            }
                            else if (arrFT_CwithBCProject[Y, 10]?.ToString() == loomName)
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

                        loomSortList.Distinct().ToList();

                        // if group number start with a letters
                        var filteredLoomSortList = loomSortList
                      .Where(x => !Regex.IsMatch(x.Group, @"^[A-Za-z]")) // Exclude groups starting with a letter
                        .ToList();

                        loomSortList = filteredLoomSortList
                      .OrderBy(x =>
                      {
                          // Extract main number from the beginning of the string
                          var match = Regex.Match(x.Group, @"^(\d+)");
                          return match.Success ? int.Parse(match.Groups[1].Value) : int.MaxValue;
                      })
                        .ThenBy(x =>
                        {
                            // Extract alphabetical part after the number (case-insensitive)
                            var match = Regex.Match(x.Group, @"^(\d+)([A-Za-z]*)");
                            return match.Success ? match.Groups[2].Value.ToLower() : "";
                        })
                     .ThenBy(x =>
                     {
                         // Extract sub-number after _ or - (default to 0 if not found)
                         var parts = Regex.Split(x.Group.Replace('_', '-'), "[-_]");
                         return parts.Length > 1 && int.TryParse(parts[1], out int subNumber) ? subNumber : 0;
                     })
                        .ThenBy(x => x.WireCode)  // Sort by WireCode last
                      .ToList();

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
                        int NoLoomSheets = modExcelInst.GetNumberOfSheetsRequired(loomRowCount, LoomSheetRowRequired);

                        //implementing Merge and demege.
                        int serialNo = 1;
                        int loomMergeEndRow = 0;
                        int loomMergeStartRow;
                        int loomStartRowCount = 1;
                        int loomMergeEndRowCount = 1;
                        int listCount = 0;

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

                            int sameloomListCount = loomSortList.Where(e => e.WireName == loomMerge.WireName).ToList().Count;

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
                        modExcelInst.GenerateHALReportFormat_CableList(arrFinalReportLOOM, loomName, NoLoomSheets, LoomSheetRowRequired, bundleFolderPath);
                        arrCustomLoom = new object[arrFT_CwithBC.GetUpperBound(0), 9];
                        Logging.Info($"{loomName} Loom report Completed");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"{loomName}:incorrect cable group id");
                        Logging.Error($"{loomName}: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Logging.Error(ex.Message);
                MessageBox.Show("ModMain ## 07" + ex.Message);
            }
        }

        //this method return NERD value based on layerNumber
        public static string layerAssign(string layerNuber)
        {
            string layerChar;

            switch (layerNuber)
            {
                case "158":
                    layerChar = "D";
                    break;
                case "152":
                    layerChar = "R";
                    break;
                case "160":
                case "154":
                    layerChar = "N";
                    break;
                case "169":
                    layerChar = "E";
                    break;
                default:
                    layerChar = "";
                    break;
            }
            return layerChar;

        }

        // this method generates Excel reports for Megger
        public void MeggerSchedulesheet(modExcel modExcelInst, List<string> selectedLoomlistforMegger, string reportsFolderPath, List<ElectreObject> filteredElecCollection)
        {
            string[,] arrContList = null;

            var partnumbers = partnumbersList.ToList();

            // var selectedLoomlistforMegger = _customReportWindow.SelectedMeggerLoomList;

            var filteredData = filteredElecCollection.Where(e => selectedLoomlistforMegger.Contains(e.BundleName) && string.IsNullOrEmpty(e.NoMegger)).ToList();
           // filteredData = filteredData.Where(e => string.IsNullOrEmpty(e.NoMegger)).ToList();            

            filteredData = filteredData.OrderByDescending(obj => obj.ComponentType == "EQU") // "EQU" first
                                                  .ThenBy(obj =>
                                                      modMain.connectorMap.Keys.Any(suffix =>
                                                          obj.ConnectorName.EndsWith($"_{suffix}", StringComparison.OrdinalIgnoreCase)
                                                      ) ? 1 : 0 // if it matches a known connector suffix → 1 (comes later), else → 0 (comes first)
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
                DeleteExistingFiles(fullPathMeggerFolder);
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

                        object[,] arrCustomLoom_CC = ConvertDictionaryTo2DArray(sortedMatchedComponents);
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
                        string[,] arrNoMegger = modExcelInst.convertNoMeggerObjto2dArray(ElecCollectionforNoMegger); // 
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
                        lowMeggerData = Convert2DArrayToList(arrContList); // populating lowMeggerData with Continuity data
                        // filtering the filteredData based on only the Continuity (Source and destination) components
                        filteredData = filteredData.Where(e => arrListOfContinuityComponents.Contains(e.ConnectorName)).ToList();
                        MeggerSheet5(filteredData, reportsFolderPath);

                        // Add hyperlinks to Megger_*.txt
                        AddHyperlinksToSheet(sheet, fullPathMeggerFolder, "Megger_");
                    }
                    else if (proceedWithRemainingSheets && sheet.Name == "High Megger")
                    {
                        MeggerSheet6and7("High", reportsFolderPath);
                        // ✅ Add hyperlinks to HighMegger_*.txt
                        AddHyperlinksToSheet(sheet, fullPathMeggerFolder, "HighMegger_");
                    }
                    else if (proceedWithRemainingSheets && sheet.Name == "Low Megger")
                    {
                        MeggerSheet6and7("Low", reportsFolderPath);
                        // ✅ Add hyperlinks to LowMegger_*.txt
                        AddHyperlinksToSheet(sheet, fullPathMeggerFolder, "LowMegger_");
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

        private void AddHyperlinksToSheet(Worksheet sheet, string folderPath, string filePrefix)
        {
            string[] filePaths = Directory.GetFiles(folderPath, $"{filePrefix}*.txt");
            int row = 1;

            foreach (string file in filePaths.OrderBy(f => f))
            {
                string fileName = Path.GetFileName(file);
                string displayText = $"Refer to {fileName} - Click to Open";

                Range hyperlinkCell = sheet.Cells[row, 1]; // Column A
                sheet.Columns[1].ColumnWidth = 40;

                sheet.Hyperlinks.Add(hyperlinkCell, file, Type.Missing, Type.Missing, displayText);
                hyperlinkCell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                hyperlinkCell.Font.Bold = true;

                row++;
            }
        }

        public static List<string> Convert2DArrayToList(string[,] arrContList)
        {
            var result = new List<string>();

            int rows = arrContList.GetLength(0);

            for (int i = 0; i < rows; i++)
            {
                string connector1 = arrContList[i, 0];
                string pin1 = arrContList[i, 1];
                string connector2 = arrContList[i, 2];
                string pin2 = arrContList[i, 3];
                if (string.IsNullOrWhiteSpace(connector1) ||
                    string.IsNullOrWhiteSpace(pin1) ||
                    string.IsNullOrWhiteSpace(connector2) ||
                    string.IsNullOrWhiteSpace(pin2))
                {
                    continue;
                }
                string resultLine = $"{connector1};{pin1};{connector2};{pin2};Low Megger";
                result.Add(resultLine);
            }

            return result;
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

                          arrCustomLoom = ConvertDictionaryTo2DArray(matchedComponents);
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

                foreach (string partNumber in partnumbers)
                {
                    if (!ElecList.Any(e => e.PartNumber == partNumber))
                        continue;

                    var connectorNames = partNumbersandConnectors[partNumber];

                    foreach (var connectorName in connectorNames)
                    {
                        // Match only connectors listed in continuity components
                        if (string.IsNullOrEmpty(connectorName) || !arrListOfContinuityComponents.Contains(connectorName))
                            continue;

                        // Retrive the Gauge from data extraction based on the matched ConnectorName
                        var connectorGauge_dataextraction = ElecCollection.Where(e => e.ConnectorName == connectorName).Select(e => e.Gauge).FirstOrDefault();

                        // Retrive the gauge of the Part Number from the Library 
                        var gauge_fromLibrary = ElecList.Where(x => x.PartNumber == partNumber && x.Gauge == connectorGauge_dataextraction)
                                                         .Select(x => x.Gauge)
                                                         .FirstOrDefault();

                        // If both dataextraction and library Gauges are matched then retrive the maxSizeofConnector
                        maxsizofconnector = ElecList.Where(e => e.PartNumber == partNumber && e.Gauge == gauge_fromLibrary)
                                                    .Select(e => e.MaxsizeofPins)
                                                    .Select(p => Convert.ToInt32(p))
                                                    .FirstOrDefault();

                        // Retrive the pins List of the matched Part number, MaxsizeofPins and Gauge
                        pinlistfromliabarary = ElecList
                            .Where(x => x.PartNumber == partNumber
                                && x.MaxsizeofPins == maxsizofconnector.ToString()
                                && x.Gauge == gauge_fromLibrary)
                            .Select(x => x.Pinnumber)
                            .ToList();

                        // Looping through each set of Pin List and adding it into allconnectorsandpins object
                        foreach (string pin in pinlistfromliabarary)
                        {
                            if (pin.Contains("-"))
                            {
                                string[] parts = pin.Split('-');
                                string start = parts[0];
                                string end = parts[1];

                                // Numeric range (1-20)
                                if (int.TryParse(start, out int startNum) && int.TryParse(end, out int endNum))
                                {
                                    for (int num = startNum; num <= endNum; num++)
                                    {
                                        allconnectorsandpins.Add(connectorName);
                                        allconnectorsandpins.Add(num.ToString());
                                    }
                                }
                                // Alphabetic range ( A-Z )
                                else if (start.Length == 1 && end.Length == 1 && char.IsLetter(start[0]) && char.IsLetter(end[0]))
                                {                                   
                                    for (char c = start[0]; c <= end[0]; c++) 
                                    { 
                                        allconnectorsandpins.Add(connectorName);
                                        allconnectorsandpins.Add(c.ToString());
                                    }
                                }
                                // Alphanumeric range with multiple letters prefix and number ranges (e.g., "A1-A5", "AA10-AA15", "AAA9-AAA19")
                                else if (start.Length > 1 && end.Length > 1 && start.TakeWhile(char.IsLetter).Count() == end.TakeWhile(char.IsLetter).Count() &&
                                         start.TakeWhile(char.IsLetter).All(char.IsLetter) &&
                                         int.TryParse(start.Substring(start.TakeWhile(char.IsLetter).Count()), out int startNumRange) &&
                                         int.TryParse(end.Substring(end.TakeWhile(char.IsLetter).Count()), out int endNumRange) &&
                                         start.Substring(0, start.TakeWhile(char.IsLetter).Count()) == end.Substring(0, start.TakeWhile(char.IsLetter).Count()))
                                {
                                    // Extract the letter prefix (e.g., "A", "AA", "AAA")
                                    string letterPrefix = start.Substring(0, start.TakeWhile(char.IsLetter).Count());

                                    // Extract numeric parts of the start and end pins (e.g., 1 from "A1", 5 from "A5")
                                    int startNumber = int.Parse(start.Substring(start.TakeWhile(char.IsLetter).Count()));
                                    int endNumber = int.Parse(end.Substring(end.TakeWhile(char.IsLetter).Count()));

                                    // Generate the range for the numbers
                                    for (int num = startNumber; num <= endNumber; num++)
                                    {
                                        allconnectorsandpins.Add(connectorName);
                                        allconnectorsandpins.Add(letterPrefix + num.ToString());  // Combine the prefix and number
                                    }
                                }
                                // Numeric range with multi-letter suffix (e.g., "2A-5A", "23AA-25AA", "227AAA-2343AAA")
                                else if (start.Length > 1 && end.Length > 1 &&
                                         int.TryParse(start.Substring(0, start.Length - start.TakeWhile(char.IsLetter).Count()), out int startNum4) &&
                                         int.TryParse(end.Substring(0, end.Length - end.TakeWhile(char.IsLetter).Count()), out int endNum4) &&
                                         start.Substring(start.Length - start.TakeWhile(char.IsLetter).Count()) == end.Substring(end.Length - end.TakeWhile(char.IsLetter).Count()))
                                {
                                    // Extract the numeric part and letter suffix part
                                    string letterSuffix = start.Substring(start.Length - start.TakeWhile(char.IsLetter).Count());  // Extract the letter suffix (e.g., "AA")
                                    int startNumber = int.Parse(start.Substring(0, start.Length - letterSuffix.Length));  // Extract the number part before the suffix
                                    int endNumber = int.Parse(end.Substring(0, end.Length - letterSuffix.Length));  // Extract the number part before the suffix

                                    // Loop through the number range and add the corresponding pin names
                                    for (int num = startNumber; num <= endNumber; num++)
                                    {
                                        allconnectorsandpins.Add(connectorName);
                                        allconnectorsandpins.Add(num.ToString() + letterSuffix);  // Combine the number and the letter suffix
                                    }
                                }

                                else if (start.Equals(end, StringComparison.OrdinalIgnoreCase))
                                {
                                    allconnectorsandpins.Add(connectorName);
                                    allconnectorsandpins.Add(start);  // Just add the single pin
                                }
                                else
                                {
                                    allconnectorsandpins.Add(connectorName);
                                    allconnectorsandpins.Add(pin);
                                }
                            }
                            else
                            {
                                allconnectorsandpins.Add(connectorName);
                                allconnectorsandpins.Add(pin);
                            }
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

                object[,] arrCustomLoom_PinList_sorted = new object[Rows, cols];


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

        // more advanced method using partioner and temp files
        public void MeggerSheet5(List<ElectreObject> filteredData, string reportsFolderPath)
        {
            try
            {
                string templpath = Path.Combine(GlobalVar.StrtCmd, "templ");
                string folderpath = Path.Combine(templpath, reportsFolderPath);
               // string finalFilePath = Path.Combine(folderpath, "Megger.txt");

                if (!Directory.Exists(folderpath))
                    Directory.CreateDirectory(folderpath);

                highMeggerData = new ConcurrentBag<string>();
                var tempFiles = new ConcurrentBag<string>();

                int count = filteredData.Count;
                if (count == 0)
                {
                    return;
                }

                var redundancySet = new ConcurrentDictionary<string, byte>(Environment.ProcessorCount * 2, count * 2);
                int partitionSize = Math.Max(200,count / Environment.ProcessorCount);

                var rangePartitioner = Partitioner.Create(0, count, partitionSize);

                Parallel.ForEach(rangePartitioner, new ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount }, range =>
                {
                    for (int i = range.Item1; i < range.Item2; i++)
                    {
                        var E1 = filteredData[i];

                        for (int j = 0; j < count; j++)
                        {
                            if (i == j) continue;
                            var n1 = filteredData[j];

                            // Ensure symmetric key to avoid A-B and B-A duplicates
                            string partA = $"{E1.ConnectorName}#{E1.PinNumber}";
                            string partB = $"{n1.ConnectorName}#{n1.PinNumber}";
                            string key = string.Compare(partA, partB) < 0
                                ? $"{partA}##{partB}"
                                : $"{partB}##{partA}";
                                
                            if (!redundancySet.TryAdd(key, 0))
                                continue;

                            string status = (string.Equals(n1.WireNumber, E1.WireNumber) &&
                                             string.Equals(n1.SubNet, E1.SubNet))
                                             ? "Low Megger"
                                             : "High Megger";

                            string resultLine = string.Empty;
                            if (status == "High Megger")
                            {
                                string connectionLine = $"{E1.ConnectorName};{E1.PinNumber};{n1.ConnectorName};{n1.PinNumber}";
                                //if connectionline is exists in lowMegger then avoid it in high megger
                                bool existsInLowMegger = lowMeggerData.Any(line => line.StartsWith($"{connectionLine};"));
                                if (!existsInLowMegger)
                                {
                                    resultLine = $"{connectionLine};{status}";
                                    highMeggerData.Add(resultLine);
                                }                                
                            }
                        }
                    }
                });

                if (listConnectorPinsLibrary.Count > 0)
                {
                    // Partition the source list for parallel processing
                    var rangePartitioner1 = Partitioner.Create(0, listConnectorPinsLibrary.Count);

                    Parallel.ForEach(rangePartitioner1, new ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount }, range =>
                    {
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

                                if (string.Equals(srcConnector, tgtConnector) &&
                                    string.Equals(srcPin, tgtPin))
                                    continue;

                                // Create symmetric key
                                string keyA = $"{srcConnector}#{srcPin}";
                                string keyB = $"{tgtConnector}#{tgtPin}";
                                string pairKey = string.Compare(keyA, keyB) < 0
                                    ? $"{keyA}##{keyB}"
                                    : $"{keyB}##{keyA}";

                                if (!redundancySet.TryAdd(pairKey, 0))
                                    continue;

                                string line = $"{srcConnector};{srcPin};{tgtConnector};{tgtPin};High Megger";
                                highMeggerData.Add(line);
                            }

                            // Loop through pin list itself
                            foreach(var target in listConnectorPinsLibrary)
                            {
                                string tgtConnector = target.ConnectorName;
                                string tgtPin = target.PinNumber;

                                if (string.Equals(srcConnector, tgtConnector) &&
                                    string.Equals(srcPin, tgtPin))
                                    continue;

                                // Create symmetric key
                                string keyA = $"{srcConnector}#{srcPin}";
                                string keyB = $"{tgtConnector}#{tgtPin}";
                                string pairKey = string.Compare(keyA, keyB) < 0
                                    ? $"{keyA}##{keyB}"
                                    : $"{keyB}##{keyA}";

                                if (!redundancySet.TryAdd(pairKey, 0))
                                    continue;

                                string line = $"{srcConnector};{srcPin};{tgtConnector};{tgtPin};High Megger";
                                highMeggerData.Add(line);
                            }
                        }
                    });
                }

                // Final Write to File (Streaming - avoids large in-memory strings)

                // Sort HighMeggerData by FromConnector and FromPin
                highMeggerData = new ConcurrentBag<string>(
                    highMeggerData
                        .OrderByDescending(line =>
                        {
                            var parts = line.Split(';');
                            return parts.Length > 1 ? parts[0] : string.Empty; // FromConnector
                        })
                        .ThenByDescending(line =>
                        {
                            var parts = line.Split(';');
                            return parts.Length > 2 ? parts[1] : string.Empty; // FromPin
                        }));

                // Final Write to File (Streaming - avoids large in-memory strings)
                long maxLines = highMeggerData.Count + lowMeggerData.Count;
                long maxLinesPerSheet = 5000000; // 5 million lines per sheet
                int totalSheets = (int)Math.Ceiling((double)maxLines / maxLinesPerSheet);
                Logging.Info($"Total Megger lines to write: {maxLines}");
                Logging.Info($"Total Megger sheets to create: {totalSheets}");

                int sheetIndex = 1;
                long lineCounterInCurrentSheet = 0;

                IEnumerable<string> combinedData = lowMeggerData.Concat(highMeggerData); // No ToList()
                StreamWriter writer = null;

                try
                {
                    writer = new StreamWriter(Path.Combine(folderpath, $"Megger_{sheetIndex}.txt"), false, Encoding.UTF8, bufferSize: 65536);
                    writer.WriteLine("From;TO");
                    writer.WriteLine("FromConnector;FromPin;ToConnector;ToPin;Megger");

                    foreach (var line in combinedData)
                    {
                        if (lineCounterInCurrentSheet >= maxLinesPerSheet)
                        {
                            Logging.Info($"Megger_{sheetIndex}.txt created successfully");
                            writer.Dispose(); // Close current file
                            sheetIndex++;
                            lineCounterInCurrentSheet = 0;

                            writer = new StreamWriter(Path.Combine(folderpath, $"Megger_{sheetIndex}.txt"), false, Encoding.UTF8, bufferSize: 65536);
                            writer.WriteLine("From;TO");
                            writer.WriteLine("FromConnector;FromPin;ToConnector;ToPin;Megger");
                        }

                        writer.WriteLine(line);
                        lineCounterInCurrentSheet++;
                    }
                }
                finally
                {
                    writer?.Dispose(); // Ensure the last file is closed properly
                }

                //commented on July 17 2025 before implementing splitting the files
                /*using (var writer = new StreamWriter(finalFilePath, false, Encoding.UTF8, bufferSize: 65536))
                {
                    writer.WriteLine("From;TO");
                    writer.WriteLine("FromConnector;FromPin;ToConnector;ToPin;Megger");

                    foreach (var line in lowMeggerData)
                        writer.WriteLine(line);

                    // Sort HighMeggerData by FromConnector and FromPin
                    highMeggerData = new ConcurrentBag<string>(
                        highMeggerData
                            .OrderByDescending(line =>
                            {
                                var parts = line.Split(';');
                                return parts.Length > 1 ? parts[0] : string.Empty; // FromConnector
                            })
                            .ThenByDescending(line =>
                            {
                                var parts = line.Split(';');
                                return parts.Length > 2 ? parts[1] : string.Empty; // FromPin
                            }));


                    foreach (var line in highMeggerData)
                        writer.WriteLine(line);
                }*/

                Logging.Info("Megger sheet generation completed successfully.");
            }
            catch (Exception ex)
            {
                MessageBox.Show($" Error Megger sheet: {ex.Message}");
                Logging.Error($"Megger Sheet 5: Error occured : {ex.Message}");
            }
        }

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

                IEnumerable<string> outputLines = meggerType == "High" ? highMeggerData : lowMeggerData;

                long maxLinesPerSheet = 5_000_000;
                long lineCounterInCurrentSheet = 0;
                int sheetIndex = 1;

                StreamWriter writer = null;

                try
                {
                    string filePath = Path.Combine(folderpath, $"{meggerType}Megger_{sheetIndex}.txt");
                    writer = new StreamWriter(filePath, false, Encoding.UTF8, bufferSize: 65536);
                    writer.WriteLine("From;TO");
                    writer.WriteLine("FromConnector;FromPin;ToConnector;ToPin;Megger");

                    foreach (var line in outputLines)
                    {
                        if (lineCounterInCurrentSheet >= maxLinesPerSheet)
                        {
                            Logging.Info($"{meggerType}Megger_{sheetIndex}.txt created successfully");

                            writer.Dispose();
                            sheetIndex++;
                            lineCounterInCurrentSheet = 0;

                            filePath = Path.Combine(folderpath, $"{meggerType}Megger_{sheetIndex}.txt");
                            writer = new StreamWriter(filePath, false, Encoding.UTF8, bufferSize: 65536);
                            writer.WriteLine("From;TO");
                            writer.WriteLine("FromConnector;FromPin;ToConnector;ToPin;Megger");
                        }

                        writer.WriteLine(line);
                        lineCounterInCurrentSheet++;
                    }

                    Logging.Info($"{meggerType}Megger_{sheetIndex}.txt created successfully");
                }
                finally
                {
                    writer?.Dispose(); // Ensure last file is closed
                }

                Logging.Info($"{meggerType} Megger sheet generation completed successfully.");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error {meggerType}: " + ex.Message);
                Logging.Error($"Error occurred in {meggerType} Megger sheet: {ex.Message}");
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
               
        //this method Converts data from Dictionary to 2d array for Megger sheet3
        public static object[,] ConvertDictionaryTo2DArray(Dictionary<string, List<string>> dict)
        {
            try
            {

                int rows = dict.Count;
                int cols = dict.Values.Max(list => list.Count) + 1; // +1 for key column

                object[,] array2D = new object[rows, cols];

                int row = 0;
                foreach (var kvp in dict)
                {
                    array2D[row, 0] = kvp.Key; // First column for keys
                    for (int col = 0; col < kvp.Value.Count; col++)
                    {
                        array2D[row, col + 1] = kvp.Value[col]; // Remaining columns for values
                    }
                    row++;
                }
                return array2D;
            }
            catch (Exception ex)
            {

                // Log error message
                Logging.Error("modMain: #009 " + ex.Message);


            }
            return new object[0, 0];  // Return empty array on error
        }

        public void CableList_Report_Preprocessing(object[,] iarr, string reportName, modExcel modExcelInst)
        {
            try
            {
                int maxRows = modExcelInst.MaxRowinArray(iarr);
                int finalRowCount = 0;

                const int initialRowOffset = 1;  // Excel is 1-indexed
                const int loomInitialRow = 9;
                const int loomSheetRowLimit = 37;

                object[,] arrFinalReportLOOM = new object[maxRows, 17];
                Excel.Worksheet activeSheet = (Excel.Worksheet)modExcelInst.ExcelApp.ActiveWorkbook.ActiveSheet;

                for (int row = 1; row <= maxRows; row++)
                {
                    string wireCode = Convert.ToString((activeSheet.Cells[row + 1, 5] as Excel.Range)?.Value2);
                    string sourcePin = Convert.ToString((activeSheet.Cells[row + 1, 1] as Excel.Range)?.Value2);
                    string sourceConn = Convert.ToString((activeSheet.Cells[row + 1, 2] as Excel.Range)?.Value2);
                    string destPin = Convert.ToString((activeSheet.Cells[row + 1, 3] as Excel.Range)?.Value2);
                    string destConn = Convert.ToString((activeSheet.Cells[row + 1, 4] as Excel.Range)?.Value2);
                    string destination = Convert.ToString((activeSheet.Cells[row + 1, 7] as Excel.Range)?.Value2);
                    string remarks = Convert.ToString((activeSheet.Cells[row + 1, 6] as Excel.Range)?.Value2);

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

                int numberOfSheets = modExcelInst.GetNumberOfSheetsRequired(finalRowCount, loomSheetRowLimit);
                int currentSheet = 1;
                int serialNumber = 0;
                int startRow = 0;

                while (startRow < finalRowCount - 1)
                {
                    int baseRow = startRow;
                    if (baseRow > currentSheet * loomSheetRowLimit)
                        currentSheet++;

                    string[] baseWireParts = Convert.ToString(arrFinalReportLOOM[baseRow, 2])?.Split('/');
                    string baseWireKey = baseWireParts?.Length >= 2 ? baseWireParts[0] + baseWireParts[1] : "";

                    for (int compareRow = baseRow + 1; compareRow < finalRowCount; compareRow++)
                    {
                        string[] compareWireParts = Convert.ToString(arrFinalReportLOOM[compareRow, 2])?.Split('/');
                        string compareWireKey = compareWireParts?.Length >= 2 ? compareWireParts[0] + compareWireParts[1] : "";

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

        // This Dictionary is used for Swapping the EQU connectors
        public static Dictionary<string, string> connectorMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "J1", "a" }, { "a", "J1" },
            { "J2", "b" }, { "b", "J2" },
            { "J3", "c" }, { "c", "J3" },
            { "J4", "d" }, { "d", "J4" },
            { "J5", "e" }, { "e", "J5" },
            { "J6", "f" }, { "f", "J6" },
            { "J7", "g" }, { "g", "J7" },
            { "J8", "h" }, { "h", "J8" },
            { "J9", "j" }, { "j", "J9" },   // Skipped 'i'
            { "J10", "k" }, { "k", "J10" },
            { "J11", "l" }, { "l", "J11" },
            { "J12", "m" }, { "m", "J12" },
            { "J13", "n" }, { "n", "J13" },
            { "J14", "p" }, { "p", "J14" }, // Skipped 'o'
            { "J15", "q" }, { "q", "J15" },
            { "J16", "r" }, { "r", "J16" },
            { "J17", "s" }, { "s", "J17" },
            { "J18", "t" }, { "t", "J18" },
            { "J19", "u" }, { "u", "J19" },
            { "J20", "v" }, { "v", "J20" },
            { "J21", "w" }, { "w", "J21" },
            { "J22", "x" }, { "x", "J22" },
            { "J23", "y" }, { "y", "J23" },
            { "J24", "z" }, { "z", "J24" },
        };

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
                  int MaxARR = modExcelInst.MaxRowinArray(iarr);

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
                  int NoLoomSheets = modExcelInst.GetNumberOfSheetsRequired(c, LoomSheetRowRequired);

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
