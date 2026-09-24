using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Excel.Application;
using Electre_Customize_DotNet.Helpers.CableList;
using Electre_Customize_DotNet.Helpers.Continuity;
using Electre_Customize_DotNet.Helpers.Global;
using Electre_Customize_DotNet.Helpers.PowerOn;
using Electre_Customize_DotNet.Objects;
using Electre_Customize_DotNet.Logs;
using Electre_Customize_DotNet.Reports;
using static System.Runtime.InteropServices.JavaScript.JSType;
using Electre_Customize_DotNet.ReportUI;
using System.Configuration;
using System.Diagnostics;
using System.IO;

namespace Electre_Customize_DotNet.MainOperation
{
    internal class modExcel
    {
        public Application ExcelApp;
        public Workbook ExcelWB;
        public Workbook TempWB;
        public Worksheet ExcelSheet;
        public Worksheet WSheet;
        public int ExcelRow;
        public Workbook ReportWB;
        public Worksheet ReportWS;
        public Worksheet ReportWS1;
        public Worksheet ReportWS2;
        public Worksheet ReportWS3;
        public Worksheet ReportWS4;
        public Worksheet ReportWS5; // for High Megger
        public Worksheet ReportWS6; // for Low Megger
        public int ReportRow;
        public int ReportAppendRow;
        public int ReportAppendRow1;
        public int ReportAppendRow2;
        public int ReportAppendRow3;
        public int ReportAppendRow4;
        public int ReportAppendRow5;  // for High Megger
        public int ReportAppendRow6;  // for Low Megger
        public int LoomInitialRow, LoomInitialCol;
        public string sCOMPONENT_CATALOG_File;
        public string OOTBPaneltotalWeight;
        public static Dictionary<string, List<string>> matchedComponents = new Dictionary<string, List<string>>();
        private XlCalculation _savedCalculation;
        private bool _excelSuspended;
        private List<PowerOnExcelBlock> _powerOnBlocks;

        public void InitiateExcel()
        {
            try
            {
                ExcelApp = new Application();
                ExcelApp.Visible = false;
                ExcelApp.DisplayAlerts = false;
                SuspendExcelUpdates();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error initiating Excel:" + ex.Message);
                Logging.Error("Error initiating Excel: " + ex.Message);

            }
        }

        private void SuspendExcelUpdates()
        {
            if (ExcelApp == null || _excelSuspended)
                return;

            try
            {
                _savedCalculation = ExcelApp.Calculation;
            }
            catch
            {
                _savedCalculation = XlCalculation.xlCalculationAutomatic;
            }

            ExcelApp.ScreenUpdating = false;
            ExcelApp.EnableEvents = false;
            ExcelApp.DisplayAlerts = false;
            ExcelApp.Calculation = XlCalculation.xlCalculationManual;
            try { ExcelApp.PrintCommunication = false; } catch { }
            try { ExcelApp.Interactive = false; } catch { }
            try { ExcelApp.AskToUpdateLinks = false; } catch { }
            try { ExcelApp.EnableAnimations = false; } catch { }
            _excelSuspended = true;
        }

        private void ResumeExcelUpdates()
        {
            if (ExcelApp == null || !_excelSuspended)
                return;

            try { ExcelApp.PrintCommunication = true; } catch { }
            try { ExcelApp.Calculation = _savedCalculation; } catch { }
            try { ExcelApp.EnableAnimations = true; } catch { }
            try { ExcelApp.Interactive = true; } catch { }
            ExcelApp.EnableEvents = true;
            ExcelApp.ScreenUpdating = true;
            _excelSuspended = false;
        }

        private static void ReleaseCom(object comObj)
        {
            ExcelRangeHelper.ReleaseCom(comObj);
        }

        private void DiscardReportWorkbook(bool save = false)
        {
            if (ReportWB == null)
                return;
            try { ReportWB.Close(save); } catch { }
            ReleaseCom(ReportWS);
            ReleaseCom(ReportWB);
            ReportWS = null;
            ReportWB = null;
        }

        private static string ToA1(int row, int col)
        {
            return ExcelRangeHelper.ToA1(row, col);
        }

        private Worksheet? FindSheetByName(Workbook workbook, string sheetName, bool ignoreCase = true)
        {
            return ExcelSheetOps.FindByName(workbook, sheetName, ignoreCase);
        }

        private void WriteBlock(Worksheet ws, int startRow, int startCol, object[,] block)
        {
            ExcelSheetOps.WriteBlock(ws, startRow, startCol, block);
        }

        private void UngroupSheets()
        {
            ExcelSheetOps.Ungroup(ReportWB);
        }

        private void SelectSheetGroup(int fromSheet, int toSheet)
        {
            ExcelSheetOps.SelectGroup(ReportWB, fromSheet, toSheet);
        }

        private void ClearCutCopyMode()
        {
            ExcelSheetOps.ClearCutCopyMode(ExcelApp);
        }

        // this method creates Excel Workbook for megger
        public string CreateNewWorkbookforMeggerScheduler(string iWorkbookName, string iSheetName1, string iSheetName2, string iSheetName3, string iSheetName4, string iSheetName5, string iSheetName6, string iSheetName7, string reportType)
        {
            try
            {
                SuspendExcelUpdates();
                ExcelApp.DisplayAlerts = false;
                DiscardReportWorkbook();

                string[] sheetNames =
                {
                    iSheetName1, iSheetName2, iSheetName3, iSheetName4,
                    iSheetName5, iSheetName6, iSheetName7
                };

                try
                {
                    ReportWB = ExcelApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
                    ((Worksheet)ReportWB.Sheets[1]).Name = sheetNames[0];
                    for (int i = 1; i < sheetNames.Length; i++)
                    {
                        Worksheet added = (Worksheet)ReportWB.Sheets.Add(After: ReportWB.Sheets[ReportWB.Sheets.Count]);
                        added.Name = sheetNames[i];
                    }
                }
                catch
                {
                    ReportWB = ExcelApp.Workbooks.Add();
                    for (int i = 0; i < sheetNames.Length; i++)
                    {
                        Worksheet added = (Worksheet)ReportWB.Sheets.Add(After: ReportWB.Sheets[ReportWB.Sheets.Count]);
                        added.Name = sheetNames[i];
                    }
                    DeleteDefaultSheets();
                }

                // Save the workbook
                string reportPath = Path.Combine(GlobalVar.ReportFolderGlobal, reportType);

                if (!Directory.Exists(reportPath))
                {
                    Directory.CreateDirectory(reportPath);
                }

                string workbookPath = Path.Combine(reportPath, iWorkbookName);
                ReportWB.SaveAs(workbookPath, XlFileFormat.xlExcel8);
                if (!File.Exists($"{workbookPath}.xls"))
                {
                    return "null";
                };

                Logging.Info($"megger empty Report {workbookPath}.xls");
                return workbookPath + ".xls";

            }
            catch (COMException ex)
            {
                MessageBox.Show("Excel ##06: " + ex.Message);
                Logging.Info("Error interacting with Excel: " + ex.Message);
                return "null";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Excel ##07: " + ex.Message);
                Logging.Info("An error occurred: " + ex.Message);
                return "null";
            }
        }

        /*   public string CreateNewWorkbookforMeggerScheduler(string iWorkbookName, string iSheetName1, string iSheetName2, string iSheetName3, string iSheetName4, string reportType)
           {
               try
               {
                   ReportWB = ExcelApp.Workbooks.Add();
                   ExcelApp.DisplayAlerts = false;




                   // Add a new sheet and name it
                   ReportWB.Sheets.Add(After: ReportWB.Sheets[ReportWB.Sheets.Count]);
                   Worksheet newSheet1 = (Worksheet)ReportWB.Sheets[ReportWB.Sheets.Count];
                   newSheet1.Name = iSheetName1;

                   ReportWB.Sheets.Add(After: ReportWB.Sheets[ReportWB.Sheets.Count]);
                   Worksheet newSheet2 = (Worksheet)ReportWB.Sheets[ReportWB.Sheets.Count];
                   newSheet2.Name = iSheetName2;

                   ReportWB.Sheets.Add(After: ReportWB.Sheets[ReportWB.Sheets.Count]);
                   Worksheet newSheet3 = (Worksheet)ReportWB.Sheets[ReportWB.Sheets.Count];
                   newSheet3.Name = iSheetName3;

                   ReportWB.Sheets.Add(After: ReportWB.Sheets[ReportWB.Sheets.Count]);
                   Worksheet newSheet4 = (Worksheet)ReportWB.Sheets[ReportWB.Sheets.Count];
                   newSheet4.Name = iSheetName4;

                   // Delete default sheets
                   DeleteDefaultSheets();

                   // Save the workbook
                   string reportPath = Path.Combine(GlobalVar.ReportFolderGlobal, reportType);

                   if (!Directory.Exists(reportPath))
                   {
                       Directory.CreateDirectory(reportPath);
                   }

                   string workbookPath = Path.Combine(reportPath, iWorkbookName);
                   ReportWB.SaveAs(workbookPath, XlFileFormat.xlExcel8);
                   if (!File.Exists($"{workbookPath}.xls"))
                   {
                       return "null";
                   };

                   Logging.Info($"megger empty Report {workbookPath}.xls");
                   return workbookPath + ".xls";

               }
               catch (COMException ex)
               {
                   MessageBox.Show("Excel ##06: " + ex.Message);
                   Logging.Info("Error interacting with Excel: " + ex.Message);
                   return "null";
               }
               catch (Exception ex)
               {
                   MessageBox.Show("Excel ##07: " + ex.Message);
                   Logging.Info("An error occurred: " + ex.Message);
                   return "null";
               }
           }*/
        // this method creates Headres for megger Excel sheets
        public bool CreateMeggerSchedulerHeader(string workbookPath)
        {
            ReportAppendRow = 1;
            ReportAppendRow1 = 1;
            ReportAppendRow2 = 1;
            ReportAppendRow3 = 1;
            ReportAppendRow4 = 1;
            ReportAppendRow5 = 1;
            ReportAppendRow6 = 1;
            try
            {
                SuspendExcelUpdates();
                if (!IsWorkbookOpen(workbookPath))
                    ReportWB = ExcelApp.Workbooks.Open(workbookPath);

                foreach (Worksheet sheet in ReportWB.Sheets)
                {
                    if (sheet.Name == "Continuity Components")
                    {

                        ReportWS = sheet as Worksheet; // Access the worksheet

                        if (ReportWS == null)
                        {
                            Logging.Error("Failed to access the active sheet.");
                            return false;

                        }

                        object[,] ccHeader = new object[1, 2];
                        ccHeader[0, 0] = "Component Name";
                        ccHeader[0, 1] = "Part Numbers";
                        WriteBlock(ReportWS, ReportAppendRow, 1, ccHeader);
                        Range headerRange = ReportWS.Range["A1", "B1"];
                        headerRange.Font.Bold = true;
                        ReleaseCom(headerRange);

                        ((Range)ReportWS.Columns["A"]).ColumnWidth = 20;
                        ((Range)ReportWS.Columns["B"]).ColumnWidth = 20;

                        ReportAppendRow++;


                    }
                    else if (sheet.Name == "Pin List")
                    {
                        ReportWS1 = sheet as Worksheet; // Access the worksheet
                                                        // ReportWS = ReportWB.ActiveSheet as Worksheet;

                        if (ReportWS1 == null)
                        {
                            Logging.Error("Failed to access the active sheet.");
                            return false;
                        }
                        object[,] pinHeader = new object[1, 2];
                        pinHeader[0, 0] = "Connectors";
                        pinHeader[0, 1] = "Pins";
                        WriteBlock(ReportWS1, ReportAppendRow1, 1, pinHeader);

                        Range headerRange = ReportWS1.Range["A1", "B1"];
                        headerRange.Font.Bold = true;
                        ReleaseCom(headerRange);

                        ((Range)ReportWS1.Columns["A"]).ColumnWidth = 20;
                        ((Range)ReportWS1.Columns["B"]).ColumnWidth = 15;
                        ReportAppendRow1++;

                    }
                    else if (sheet.Name == "Connection List")
                    {

                        ReportWS2 = sheet as Worksheet; // Access the worksheet
                                                        // ReportWS = ReportWB.ActiveSheet as Worksheet;

                        if (ReportWS2 == null)
                        {
                            Logging.Error("Failed to access the active sheet.");
                            return false;
                        }

                        object[,] fromTo = new object[1, 4];
                        fromTo[0, 0] = "From";
                        fromTo[0, 2] = "To";
                        WriteBlock(ReportWS2, ReportAppendRow2, 1, fromTo);

                        Range fromRange = ReportWS2.Range["A" + ReportAppendRow2, "B" + ReportAppendRow2];
                        Range toRange = ReportWS2.Range["C" + ReportAppendRow2, "D" + ReportAppendRow2];
                        fromRange.Merge();
                        toRange.Merge();

                        fromRange.Font.Bold = true;
                        fromRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        fromRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                        toRange.Font.Bold = true;
                        toRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        toRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                        ReleaseCom(fromRange);
                        ReleaseCom(toRange);

                        ReportAppendRow2++;

                        object[,] subHeader = new object[1, 4];
                        subHeader[0, 0] = "Connector No";
                        subHeader[0, 1] = "Pin No";
                        subHeader[0, 2] = "Connector No";
                        subHeader[0, 3] = "Pin No";
                        WriteBlock(ReportWS2, ReportAppendRow2, 1, subHeader);

                        Range headerRange = ReportWS2.Range["A2", "D2"];
                        headerRange.Font.Bold = true;
                        ReleaseCom(headerRange);

                        ((Range)ReportWS2.Columns["A"]).ColumnWidth = 15;
                        ((Range)ReportWS2.Columns["B"]).ColumnWidth = 10;
                        ((Range)ReportWS2.Columns["C"]).ColumnWidth = 15;
                        ((Range)ReportWS2.Columns["D"]).ColumnWidth = 10;

                        ReportAppendRow2++; // Move to the next row for data
                    }

                    else if (sheet.Name == "Exception")
                    {
                        ReportWS3 = sheet as Worksheet; // Access the worksheet

                        if (ReportWS3 == null)
                        {
                            Logging.Error("Failed to access the active sheet.");
                            return false;

                        }

                        object[,] exHeader = new object[1, 3];
                        exHeader[0, 0] = "SN";
                        exHeader[0, 1] = "Connectors";
                        exHeader[0, 2] = "Pins";
                        WriteBlock(ReportWS3, ReportAppendRow3, 1, exHeader);

                        Range headerRange = ReportWS3.Range["A1", "C1"];
                        headerRange.Font.Bold = true;
                        ReleaseCom(headerRange);

                        ((Range)ReportWS3.Columns["A"]).ColumnWidth = 10;
                        ((Range)ReportWS3.Columns["B"]).ColumnWidth = 20;
                        ((Range)ReportWS3.Columns["C"]).ColumnWidth = 15;

                        ReportAppendRow3++;


                    }
                    else if (sheet.Name == "Megger")
                    {
                        ReportWS4 = sheet as Worksheet; // Access the worksheet

                        if (ReportWS4 == null)
                        {
                            Logging.Error("Failed to access the active sheet.");
                            return false;
                        }

                        string templpath;
                        string folderpath;
                        string filePath;

                        templpath = Path.Combine(GlobalVar.StrtCmd, "templ");
                        folderpath = Path.Combine(templpath, "MeggerSchedulerReports");
                        filePath = Path.Combine(folderpath, "Megger.txt");


                       /* string displayText = "Refer to Megger_1.txt - Click to Open";

                        // Add hyperlink to the cell, e.g., in column D of the current ReportAppendRow4
                        Range hyperlinkCell = ReportWS4.Cells[ReportAppendRow4, 1]; // Column A
                        ReportWS4.Columns["A"].ColumnWidth = 40;
                        ReportWS4.Hyperlinks.Add(hyperlinkCell, filePath, Type.Missing, Type.Missing, displayText);

                        // Optional: highlight the cell
                        hyperlinkCell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        hyperlinkCell.Font.Bold = true;*/

                        ReportAppendRow4++;
                    }
                    else if (sheet.Name == "High Megger")
                    {

                        ReportWS5 = sheet as Worksheet; // Access the worksheet

                        if (ReportWS5 == null)
                        {
                            Logging.Error("Failed to access the active sheet.");
                            return false;
                        }

                        string templpath;
                        string folderpath;
                        string filePath;

                        templpath = Path.Combine(GlobalVar.StrtCmd, "templ");
                        folderpath = Path.Combine(templpath, "MeggerSchedulerReports");
                        filePath = Path.Combine(folderpath, "HighMegger.txt");


                      /*  string displayText = "Refer to HighMegger_1.txt - Click to Open";

                        // Add hyperlink to the cell, e.g., in column D of the current ReportAppendRow4
                        Range hyperlinkCell = ReportWS5.Cells[ReportAppendRow5, 1]; // Column A
                        ReportWS5.Columns["A"].ColumnWidth = 40;
                        ReportWS5.Hyperlinks.Add(hyperlinkCell, filePath, Type.Missing, Type.Missing, displayText);

                        // Optional: highlight the cell
                        hyperlinkCell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        hyperlinkCell.Font.Bold = true;*/

                        ReportAppendRow5++;
                    }
                    else if (sheet.Name == "Low Megger")
                    {
                        ReportWS6 = sheet as Worksheet; // Access the worksheet

                        if (ReportWS6 == null)
                        {
                            Logging.Error("Failed to access the active sheet.");
                            return false;
                        }

                        string templpath;
                        string folderpath;
                        string filePath;

                        templpath = Path.Combine(GlobalVar.StrtCmd, "templ");
                        folderpath = Path.Combine(templpath, "MeggerSchedulerReports");
                        filePath = Path.Combine(folderpath, "LowMegger.txt");


                       /* string displayText = "Refer to LowMegger_1.txt - Click to Open";

                        // Add hyperlink to the cell, e.g., in column D of the current ReportAppendRow4
                        Range hyperlinkCell = ReportWS6.Cells[ReportAppendRow6, 1]; // Column A
                        ReportWS6.Columns["A"].ColumnWidth = 40;
                       ReportWS6.Hyperlinks.Add(hyperlinkCell, filePath, Type.Missing, Type.Missing, displayText);

                        // Optional: highlight the cell
                        hyperlinkCell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        hyperlinkCell.Font.Bold = true;*/

                        ReportAppendRow6++;
                    }
                }

                Logging.Info($"Megger Scheduler report Header created for {workbookPath}");
                return true;


            }
            catch (Exception ex)
            {
                MessageBox.Show("Excel ##06: " + ex.Message);
                Logging.Error("Failed to create WireLess report Header: " + ex.Message);
                return false;
            }



        }

        /*  public bool CreateMeggerSchedulerHeader(string workbookPath)
          {
              ReportAppendRow = 1;
              ReportAppendRow1 = 1;
              ReportAppendRow2 = 1;
              ReportAppendRow3 = 1;

              try
              {
                  ReportWB = ExcelApp.Workbooks.Open(workbookPath);

                  foreach (Worksheet sheet in ReportWB.Sheets)
                  {
                      if (sheet.Name == "pin List")
                      {

                          ReportWS = sheet as Worksheet; // Access the worksheet
                          // ReportWS = ReportWB.ActiveSheet as Worksheet;

                          if (ReportWS == null)
                          {
                              Logging.Error("Failed to access the active sheet.");
                              return false;
                          }
                          // Set header cells
                          ReportWS.Cells[ReportAppendRow, 1].Value = "Connectors";
                          ReportWS.Cells[ReportAppendRow, 2].Value = "Pins";


                          // Format the header row
                          Range headerRange = ReportWS.Range["A1", "B1"];
                          headerRange.Font.Bold = true;

                          // Set column widths
                          ReportWS.Columns["A"].ColumnWidth = 20;
                          ReportWS.Columns["B"].ColumnWidth = 15;

                          ReportAppendRow++;


                      }
                      else if (sheet.Name == "Connection List")
                      {

                          ReportWS1 = sheet as Worksheet; // Access the worksheet
                          // ReportWS = ReportWB.ActiveSheet as Worksheet;

                          if (ReportWS1 == null)
                          {
                              Logging.Error("Failed to access the active sheet.");
                              return false;
                          }

                          // Set merged header cells
                          //ReportWS1.Cells[ReportAppendRow1, 1].Value = "From";
                          //ReportWS1.Cells[ReportAppendRow1, 4].Value = "To";

                          //Range fromRange = ReportWS1.Range["A1", "B1"];
                          //Range toRange = ReportWS1.Range["C1", "D1"];
                          //fromRange.Merge();
                          //toRange.Merge();

                          //fromRange.Font.Bold = true;
                          //toRange.Font.Bold = true;

                          //fromRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                          //toRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                          //fromRange.VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;
                          //toRange.VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;

                          // Write text in the first merged cells
                          ReportWS1.Cells[ReportAppendRow1, 1].Value = "From"; // Column A
                          ReportWS1.Cells[ReportAppendRow1, 3].Value = "To";   // Column C

                          // Merge Column A and B for "From"
                          Range fromRange = ReportWS1.Range["A" + ReportAppendRow1, "B" + ReportAppendRow1];

                          // Merge Column C and D for "To"
                          Range toRange = ReportWS1.Range["C" + ReportAppendRow1, "D" + ReportAppendRow1];

                          // Merge the cells
                          fromRange.Merge();
                          toRange.Merge();

                          // Apply formatting for From
                          fromRange.Font.Bold = true;
                          fromRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                          fromRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                          // Apply formatting for To
                          toRange.Font.Bold = true;
                          toRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                          toRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;


                          // Set subheaders below "From" and "To"
                          ReportAppendRow1++;

                          ReportWS1.Cells[ReportAppendRow1, 1].Value = "Connector No";
                          ReportWS1.Cells[ReportAppendRow1, 2].Value = "Pin No";
                          ReportWS1.Cells[ReportAppendRow1, 3].Value = "Connector No";
                          ReportWS1.Cells[ReportAppendRow1, 4].Value = "Pin No";


                          // Format the header row
                          Range headerRange = ReportWS1.Range["A2", "D2"];
                          headerRange.Font.Bold = true;

                          // Set column widths for readability

                          ReportWS1.Columns["A"].ColumnWidth = 15; // Connector No (From)
                          ReportWS1.Columns["B"].ColumnWidth = 10; // Pin No (From)
                          ReportWS1.Columns["C"].ColumnWidth = 15; // Connector No (To)
                          ReportWS1.Columns["D"].ColumnWidth = 10; // Pin No (To)


                          ReportAppendRow1++; // Move to the next row for data



                      }
                      else if (sheet.Name == "Continuity Components")
                      {
                          ReportWS2 = sheet as Worksheet; // Access the worksheet

                          if (ReportWS2 == null)
                          {
                              Logging.Error("Failed to access the active sheet.");
                              return false;

                          }

                          // Set header cells
                          ReportWS2.Cells[ReportAppendRow2, 1].Value = "Component Name";
                          ReportWS2.Cells[ReportAppendRow2, 2].Value = "Part Numbers";




                          // Format the header row
                          Range headerRange = ReportWS2.Range["A1", "B1"];
                          headerRange.Font.Bold = true;

                          // Set column widths
                          ReportWS2.Columns["A"].ColumnWidth = 20;
                          ReportWS2.Columns["B"].ColumnWidth = 20;


                          ReportAppendRow2++;



                      }
                      else if(sheet.Name == "Exception")
                      {
                          ReportWS3 = sheet as Worksheet; // Access the worksheet

                          if (ReportWS3 == null)
                          {
                              Logging.Error("Failed to access the active sheet.");
                              return false;

                          }

                          // Set header cells
                          ReportWS3.Cells[ReportAppendRow3, 1].Value = "SN"; 
                          ReportWS3.Cells[ReportAppendRow3, 2].Value = "Connectors";           
                          ReportWS3.Cells[ReportAppendRow3, 3].Value = "Pins";           




                          // Format the header row
                          Range headerRange = ReportWS3.Range["A1", "C1"];
                          headerRange.Font.Bold = true;

                          // Set column widths
                          ReportWS3.Columns["A"].ColumnWidth = 10;
                          ReportWS3.Columns["B"].ColumnWidth = 20;
                          ReportWS3.Columns["C"].ColumnWidth = 15;


                          ReportAppendRow3++;


                      }




                  }

                  Logging.Info($"Megger SCheduler  report Header created for {workbookPath}");
                  return true;


              }
              catch (Exception ex)
              {
                  MessageBox.Show("Excel ##06: " + ex.Message);
                  Logging.Error("Failed to create WireLess report Header: " + ex.Message);
                  return false;
              }



          }*/
        // this method delete  default sheets before creating Excel sheets
        public void DeleteDefaultSheets()
        {
            try
            {
                foreach (Worksheet sheet in ReportWB.Sheets)
                {
                    if (sheet.Name == "Sheet1" || sheet.Name == "Sheet2" || sheet.Name == "Sheet3")
                    {
                        sheet.Delete();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Excel ##03: "+ ex.Message);
                Logging.Error(ex.Message);
            }
        }
        // this method creates workbook for sheets,project wirelist, with out HAL loom reports,Components specifice Breakdow
        public string CreateNewWorkbook(string iWorkbookName, string iSheetName, string reportType)
        {
            try
            {
                SuspendExcelUpdates();
                ExcelApp.DisplayAlerts = false;
                DiscardReportWorkbook();

                try
                {
                    ReportWB = ExcelApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
                }
                catch
                {
                    ReportWB = ExcelApp.Workbooks.Add();
                    ReportWB.Sheets.Add(After: ReportWB.Sheets[ReportWB.Sheets.Count]);
                    Worksheet added = (Worksheet)ReportWB.Sheets[ReportWB.Sheets.Count];
                    added.Name = iSheetName;
                    DeleteDefaultSheets();
                }

                Worksheet newSheet = (Worksheet)ReportWB.Sheets[1];
                if (newSheet.Name != iSheetName)
                    newSheet.Name = iSheetName;
                ReportWS = newSheet;

                // Save the workbook
                string reportPath = Path.Combine(GlobalVar.ReportFolderGlobal, reportType);

                if (!Directory.Exists(reportPath))
                {
                    Directory.CreateDirectory(reportPath);
                }

                string workbookPath = Path.Combine(reportPath, iWorkbookName);
                ReportWB.SaveAs(workbookPath, XlFileFormat.xlExcel8);
                if (!File.Exists($"{workbookPath}.xls"))
                {
                    return "null";
                };

                Logging.Info($"Wirelist empty Report {workbookPath}.xls");
                return workbookPath + ".xls";

            }
            catch (COMException ex)
            {
                MessageBox.Show("Excel ##04: "+ ex.Message);
                Logging.Info("Error interacting with Excel: " + ex.Message);
                return "null";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Excel ##05: "+ ex.Message);
                Logging.Info("An error occurred: " + ex.Message);
                return "null";
            }

        }
        // this method creates header for with out HAL template loom reports
        public bool CreateloomlistReportHeader(string workbookPath)
        {
            ReportAppendRow = 1;

            try
            {
                SuspendExcelUpdates();
                if (!IsWorkbookOpen(workbookPath))
                    ReportWB = ExcelApp.Workbooks.Open(workbookPath);

                ReportWS = ReportWB.ActiveSheet as Worksheet;

                if (ReportWS == null)
                {
                    Logging.Error("Failed to access the active sheet.");
                    return false;
                }

                object[,] header = new object[1, 11];
                header[0, 0] = "FROM CONN";
                header[0, 1] = "FROM PIN";
                header[0, 2] = "TO CONN";
                header[0, 3] = "TO PIN";
                header[0, 4] = "WIRE CODE";
                header[0, 5] = "WIRE TYPE";
                header[0, 6] = "LENGTH";
                header[0, 7] = "RD";
                header[0, 8] = "LD";
                header[0, 9] = "NERD";
                header[0, 10] = "Group Number";
                WriteBlock(ReportWS, ReportAppendRow, 1, header);

                Range headerRange = ReportWS.Range["A1", "K1"];
                headerRange.Font.Bold = true;
                ReleaseCom(headerRange);

                ((Range)ReportWS.Columns["A"]).ColumnWidth = 15;
                ((Range)ReportWS.Columns["B"]).ColumnWidth = 15;
                ((Range)ReportWS.Columns["C"]).ColumnWidth = 15;
                ((Range)ReportWS.Columns["D"]).ColumnWidth = 15;
                ((Range)ReportWS.Columns["E"]).ColumnWidth = 20;
                ((Range)ReportWS.Columns["F"]).ColumnWidth = 10;
                ((Range)ReportWS.Columns["G"]).ColumnWidth = 10;
                ((Range)ReportWS.Columns["H"]).ColumnWidth = 20;
                ((Range)ReportWS.Columns["I"]).ColumnWidth = 15;
                ((Range)ReportWS.Columns["J"]).ColumnWidth = 10;
                ((Range)ReportWS.Columns["K"]).ColumnWidth = 15;
                ReportAppendRow++;

                Logging.Info($"wireList report Header created for {workbookPath}");
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Excel ##06: "+ ex.Message);
                Logging.Error("Failed to create WireLess report Header: " + ex.Message);
                return false;
            }
        }
      
        // this method creates header for Continuity Componentes 
        public bool CreateListOfContinuityComponentsReportHeader(string workbookPath)
        {
            ReportAppendRow = 1;

            try
            {
                SuspendExcelUpdates();
                if (!IsWorkbookOpen(workbookPath))
                    ReportWB = ExcelApp.Workbooks.Open(workbookPath);

                ReportWS = ReportWB.ActiveSheet as Worksheet;

                if (ReportWS == null)
                {
                    MessageBox.Show("reportWS  null");
                    Logging.Error("Failed to access the active sheet.");
                    return false;
                }

                object[,] header = new object[1, 2];
                header[0, 0] = "Component Name";
                header[0, 1] = "Part Number";
                WriteBlock(ReportWS, ReportAppendRow, 1, header);

                Range headerRange = ReportWS.Range["A1","B1"];
                headerRange.Font.Bold = true;
                ReleaseCom(headerRange);

                ((Range)ReportWS.Columns["A"]).ColumnWidth = 20;
                ((Range)ReportWS.Columns["B"]).ColumnWidth = 20;

                ReportAppendRow++;

                Logging.Info($"CreateBreakdownReportHeader created for {workbookPath}");
                return true;
            }
            catch (Exception ex)
            {
               // MessageBox.Show("Excel ##08: "+ ex.Message);
                Logging.Error("Failed to create CreateBreakdownReportHeader: " + ex.Message);
                return false;
            }
        }

        public string WriteComponentBreakdownWorkbook(string iWorkbookName, string iSheetName, string reportType, object[,] data, int dataRows)
        {
            object[,] header = new object[1, 10];
            header[0, 0] = "FROM CONN";
            header[0, 1] = "FROM PIN";
            header[0, 2] = "TO CONN";
            header[0, 3] = "TO PIN";
            header[0, 4] = "WIRE CODE";
            header[0, 5] = "WIRE TYPE";
            header[0, 6] = "LENGTH";
            header[0, 7] = "RD";
            header[0, 8] = "LD";
            header[0, 9] = " NERD";
            double[] widths = { 14.78, 14.78, 14, 16.67, 21.33, 14.11, 10, 20, 15, 10 };
            return WriteSingleWirelistFile(iWorkbookName, iSheetName, reportType, data, dataRows, header, widths, "J");
        }

        public string WriteLoomWithoutHalWorkbook(string iWorkbookName, string iSheetName, string reportType, object[,] data, int dataRows)
        {
            object[,] header = new object[1, 11];
            header[0, 0] = "FROM CONN";
            header[0, 1] = "FROM PIN";
            header[0, 2] = "TO CONN";
            header[0, 3] = "TO PIN";
            header[0, 4] = "WIRE CODE";
            header[0, 5] = "WIRE TYPE";
            header[0, 6] = "LENGTH";
            header[0, 7] = "RD";
            header[0, 8] = "LD";
            header[0, 9] = "NERD";
            header[0, 10] = "Group Number";
            double[] widths = { 15, 15, 15, 15, 20, 10, 10, 20, 15, 10, 15 };
            return WriteSingleWirelistFile(iWorkbookName, iSheetName, reportType, data, dataRows, header, widths, "K");
        }

        private string WriteSingleWirelistFile(
            string iWorkbookName,
            string iSheetName,
            string reportType,
            object[,] data,
            int dataRows,
            object[,] header,
            double[] columnWidths,
            string lastColLetter)
        {
            try
            {
                SuspendExcelUpdates();
                ExcelApp.DisplayAlerts = false;
                DiscardReportWorkbook();

                try
                {
                    ReportWB = ExcelApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
                }
                catch
                {
                    ReportWB = ExcelApp.Workbooks.Add();
                    ReportWB.Sheets.Add(After: ReportWB.Sheets[ReportWB.Sheets.Count]);
                    Worksheet added = (Worksheet)ReportWB.Sheets[ReportWB.Sheets.Count];
                    added.Name = iSheetName;
                    DeleteDefaultSheets();
                }

                Worksheet newSheet = (Worksheet)ReportWB.Sheets[1];
                if (newSheet.Name != iSheetName)
                    newSheet.Name = iSheetName;
                ReportWS = newSheet;

                ReportAppendRow = 1;
                WriteBlock(ReportWS, 1, 1, header);

                Range headerRange = ReportWS.Range["A1", lastColLetter + "1"];
                headerRange.Font.Bold = true;
                ReleaseCom(headerRange);

                string[] colLetters = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K" };
                int widthCount = Math.Min(columnWidths.Length, colLetters.Length);
                for (int i = 0; i < widthCount; i++)
                    ((Range)ReportWS.Columns[colLetters[i]]).ColumnWidth = columnWidths[i];

                if (dataRows > 0 && data != null)
                    WriteBlock(ReportWS, 2, 1, data);

                int lastRow = Math.Max(1, 1 + dataRows);
                Range dataRange = ReportWS.Range["A1", lastColLetter + lastRow];
                FormatRange(dataRange);
                ReleaseCom(dataRange);

                string reportPath = Path.Combine(GlobalVar.ReportFolderGlobal, reportType);
                if (!Directory.Exists(reportPath))
                    Directory.CreateDirectory(reportPath);

                string workbookPath = Path.Combine(reportPath, iWorkbookName);
                ReportWB.SaveAs(workbookPath, XlFileFormat.xlExcel8);
                string resultPath = workbookPath + ".xls";
                if (!File.Exists(resultPath))
                    return "null";

                Logging.Info($"Wirelist empty Report {workbookPath}.xls");
                return resultPath;
            }
            catch (COMException ex)
            {
                MessageBox.Show("Excel ##04: "+ ex.Message);
                Logging.Info("Error interacting with Excel: " + ex.Message);
                return "null";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Excel ##05: "+ ex.Message);
                Logging.Info("An error occurred: " + ex.Message);
                return "null";
            }
            finally
            {
                DiscardReportWorkbook(save: false);
            }
        }

        #region adding data to excel

        // this method Append data to EXcel sheets for megger
        public string AppendToExcelMegger(Worksheet Reportws, object[,] iarr, int iTillColumn = 0)
        {
            try
            {
                SuspendExcelUpdates();
                Logging.Info($"Started writing data to sheet: {Reportws.Name}");

                if (iarr == null || iarr.Length == 0 || iarr.GetLength(0) == 0 || iarr.GetLength(1) == 0)
                {
                    Logging.Error("Data array is null or empty.");
                    return "null";
                }

                int row = 2; // Start after headers

                if (Reportws.Name == "Continuity Components")
                {
                    var matchedComponents = Convert2DArrayToDictionary(iarr);
                    var lines = new List<object[]>();
                    var merges = new List<(int StartRow, int PartCount)>();
                    int excelRow = row;

                    foreach (var component in matchedComponents)
                    {
                        int partCount = component.Value.Count;
                        if (partCount <= 0)
                            continue;

                        int startRow = excelRow;
                        for (int p = 0; p < partCount; p++)
                        {
                            object[] line = new object[2];
                            line[0] = p == 0 ? component.Key : "";
                            line[1] = component.Value[p];
                            lines.Add(line);
                            excelRow++;
                        }
                        merges.Add((startRow, partCount));
                    }

                    if (lines.Count > 0)
                    {
                        object[,] block = new object[lines.Count, 2];
                        for (int r = 0; r < lines.Count; r++)
                        {
                            block[r, 0] = lines[r][0];
                            block[r, 1] = lines[r][1];
                        }
                        WriteBlock(Reportws, row, 1, block);

                        for (int m = 0; m < merges.Count; m++)
                        {
                            var merge = merges[m];
                            Range mergeRange = Reportws.Range[
                                ToA1(merge.StartRow, 1) + ":" + ToA1(merge.StartRow + merge.PartCount - 1, 1)];
                            mergeRange.Merge();
                            mergeRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                            mergeRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                            ReleaseCom(mergeRange);
                        }
                    }

                    ReportWB.Save();
                    Logging.Info("Data written to Continuity Components sheet.");
                    return Reportws.Name;
                }
                else if (Reportws.Name == "Pin List")
                {
                    int tillCol = iTillColumn != 0 ? iTillColumn : iarr.GetLength(1);
                    int rowsToWrite = iarr.GetLength(0);
                    if (iarr.GetLength(0) > 3)
                    {
                        for (int i = 0; i < iarr.GetLength(0); i++)
                        {
                            if (CheckIfBlankRow(iarr, i))
                            {
                                rowsToWrite = i + 1;
                                break;
                            }
                        }
                    }

                    if (rowsToWrite > 0 && tillCol > 0)
                    {
                        object[,] block = new object[rowsToWrite, tillCol];
                        for (int r = 0; r < rowsToWrite; r++)
                            for (int c = 0; c < tillCol; c++)
                                block[r, c] = iarr[r, c];
                        WriteBlock(Reportws, row, 1, block);
                        row += rowsToWrite;
                    }

                    Range range = Reportws.Range["A1", $"B{Math.Max(1, row - 1)}"];
                    FormatRange(range);
                    ReportWB.Save();

                    Logging.Info("Data written to Pin List sheet.");
                    return Reportws.Name;
                }

                else if (Reportws.Name == "Connection List")
                {
                    int tillCol = iTillColumn != 0 ? iTillColumn : iarr.GetLength(1);
                    int row1 = 3; // START from Row 3 (row 1 = merged header, row 2 = subheaders)
                    var compact = new List<object[]>();
                    for (int i = 0; i < iarr.GetLength(0); i++)
                    {
                        if (IsRowBlank(iarr, i) || CheckIfBlankRow(iarr, i))
                            continue;
                        var line = new object[tillCol];
                        for (int j = 0; j < tillCol; j++)
                            line[j] = iarr[i, j];
                        compact.Add(line);
                    }

                    if (compact.Count > 0 && tillCol > 0)
                    {
                        object[,] block = new object[compact.Count, tillCol];
                        for (int r = 0; r < compact.Count; r++)
                            for (int c = 0; c < tillCol; c++)
                                block[r, c] = compact[r][c];
                        WriteBlock(Reportws, row1, 1, block);
                        row1 += compact.Count;
                    }

                    Range range = Reportws.Range["A1", $"D{Math.Max(1, row - 1)}"];
                    FormatRange(range);
                    Excel.Range rangeToMerge = Reportws.Range["A1", "B1"];
                    rangeToMerge.Merge();
                    Excel.Range rangeToMerge1 = Reportws.Range["C1", "D1"];
                    rangeToMerge1.Merge();
                    ReportWB.Save();

                    Logging.Info("Data written to Connection List sheet.");
                    return Reportws.Name;
                }

                else if (Reportws.Name == "Exception")
                {
                    int tillCol = iTillColumn != 0 ? iTillColumn : iarr.GetLength(1);
                    int rowsToWrite = iarr.GetLength(0);
                    if (iarr.GetLength(0) > 3)
                    {
                        for (int i = 0; i < iarr.GetLength(0); i++)
                        {
                            if (CheckIfBlankRow(iarr, i))
                            {
                                rowsToWrite = i;
                                break;
                            }
                        }
                    }

                    if (rowsToWrite > 0)
                    {
                        object[,] block = new object[rowsToWrite, 3];
                        for (int i = 0; i < rowsToWrite; i++)
                        {
                            block[i, 0] = iarr[i, 0];
                            block[i, 1] = iarr[i, 1];
                            block[i, 2] = iarr[i, 2];
                        }
                        WriteBlock(Reportws, row, 1, block);
                        row += rowsToWrite;
                    }

                    Range range = Reportws.Range["A1", $"C{Math.Max(1, row - 1)}"];
                    FormatRange(range);

                    ReportWB.Save();
                    Logging.Info("Data written to Exception sheet.");
                    return Reportws.Name;
                }


                Logging.Error("No matching worksheet name.");
                return "null";
            }

            catch (Exception ex)
            {
                Logging.Error("Error in AppendToExcelMegger: " + ex.Message);
                return "null";
            }

        }

        public string[,] convertNoMeggerObjto2dArray(List<ElectreObject> noMeggerData)
        {
            if (noMeggerData == null || noMeggerData.Count == 0)
            {
                return new string[0, 0]; // Return empty array if no data
            }

            int rowCount = noMeggerData.Count;
            string[,] result = new string[rowCount, 3]; // 3 columns: Serial, ConnectorName, PinNumber

            for (int i = 0; i < rowCount; i++)
            {
                var obj = noMeggerData[i];

                result[i, 0] = (i + 1).ToString();             // Serial Number (starts from 1)
                result[i, 1] = obj.ConnectorName ?? "";        // Safe null-check
                result[i, 2] = obj.PinNumber ?? "";
            }

            return result;
        }


        /*public string AppendToExcelMegger(Worksheet Reportws, object[,] iarr, int iTillColumn = 0)
        {
            try
            {
                if (Reportws.Name == "pin List")
                {

                    int tillCol = iTillColumn != 0 ? iTillColumn : iarr.GetLength(1); // Determine the number of columns

                    for (int k = 0; k < iarr.GetLength(0); k++) // Loop through rows
                    {
                        for (int j = 0; j < tillCol; j++) // Loop through columns
                        {
                            // Populate cells in the worksheet
                            ReportWS.Cells[ReportAppendRow, j + 1].Value = iarr[k, j];
                            // Center Align Data Horizontally and Vertically
                            ReportWS.Cells[ReportAppendRow, j + 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                            ReportWS.Cells[ReportAppendRow, j + 1].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                        }
                        ReportAppendRow++; // Move to the next row in the Excel sheet

                        // Skip writing blank rows if certain conditions are met
                        if (iarr.GetLength(0) > 3) // Check if array has more than 3 rows
                        {
                            bool isBlankRow = CheckIfBlankRow(iarr, k);
                            if (isBlankRow)
                            {
                                break;
                            }
                        }
                    }

                    // Format the range in Excel
                    Range range = ReportWS.Range["A1", $"B{ReportAppendRow - 1}"];
                    FormatRange(range);


                    return ReportWS.Name;

                }
                else if (Reportws.Name == "Connection List")
                {
                    int tillCol = iTillColumn != 0 ? iTillColumn : iarr.GetLength(1); // Determine the number of columns
                    int startrow = 0;
                    if (startrow < iarr.GetLength(0))
                    {
                        startrow++;
                    }
                    for (int k = startrow; k < iarr.GetLength(0); k++) // Loop through rows
                    {
                        for (int j = 0; j < tillCol; j++) // Loop through columns
                        {
                            // Populate cells in the worksheet
                            ReportWS1.Cells[ReportAppendRow1, j + 1].Value = iarr[k, j];
                            // Center Align Data Horizontally and Vertically
                            ReportWS1.Cells[ReportAppendRow, j + 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                            ReportWS1.Cells[ReportAppendRow, j + 1].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                        }
                        ReportAppendRow1++; // Move to the next row in the Excel sheet

                        // Skip writing blank rows if certain conditions are met
                        if (iarr.GetLength(0) > 3) // Check if array has more than 3 rows
                        {
                            bool isBlankRow = CheckIfBlankRow(iarr, k);
                            if (isBlankRow)
                            {
                                break;
                            }
                        }
                    }

                    // Format the range in Excel
                    Range range = ReportWS1.Range["A1", $"D{ReportAppendRow - 1}"];
                    FormatRange(range);


                    return ReportWS1.Name;


                }
                else if(Reportws.Name == "Continuity Components")
                {
                   matchedComponents = Convert2DArrayToDictionary(iarr);

                    int row = 2;
                    // Loop through the dictionary and add component name and part number
                    foreach (var component in matchedComponents)
                    {
                        int partCount = component.Value.Count;

                        // If there are part numbers, populate the first row with the component name and merge cells
                        if (partCount > 0)
                        {
                            // Populate "Component Name" in column 1 for the first part number
                            ReportWS2.Cells[row, 1].Value = component.Key;

                            // Center align the text both horizontally and vertically for the merged cells in column 1
                            var componentCell = ReportWS2.Cells[row, 1];
                            componentCell.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                            componentCell.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;

                            // Populate "Part Number" in column 2 for the first part number
                            ReportWS2.Cells[row, 2].Value = component.Value[0];

                            // Merge the first column for the number of rows corresponding to part numbers
                            ReportWS2.Range[ReportWS2.Cells[row, 1], ReportWS2.Cells[row + partCount - 1, 1]].Merge();

                            // Loop through the remaining part numbers and populate column 2 (Part Numbers)
                            for (int i = 1; i < partCount; i++)
                            {
                                row++; // Move to the next row for the next part number

                                // Leave the first column empty (it is merged, and should remain the same for this component)
                                ReportWS2.Cells[row, 1].Value = "";

                                // Populate "Part Number" in column 2
                                ReportWS2.Cells[row, 2].Value = component.Value[i];
                            }

                            // Move to the next row after all part numbers for the current component are processed
                            row++;
                        }
                    }


                    return ReportWS2.Name;


                    ReportWB.Save();
                }

                return "null";

            }
            catch (Exception ex)
            {
                Logging.Error("An error occurred: " + ex.Message);
                return "null";
            }
            finally
            {
                if (ReportWB != null)
                {
                    if (Reportws.Name == "Continuity Components")
                    {
                        ReportWB.Close(true);
                        Marshal.ReleaseComObject(ReportWB);

                    }


                }

            }
        }*/

        // this method Append data to Excel for sheets,project wirelist, with out HAL template loom, Component specifice Breakdown reports
        public string AppendToExcel(object[,] iarr, int iTillColumn = 0)
        {
            try
            {                
                int tillCol = iTillColumn != 0 ? iTillColumn : iarr.GetLength(1);
                int rowCount = iarr.GetLength(0);
                int rowsToWrite = rowCount;

                bool compact = rowCount > 0 && !IsRowBlank(iarr, 0) && !IsRowBlank(iarr, rowCount - 1);
                if (rowCount > 3 && !compact)
                {
                    for (int k = 0; k < rowCount; k++)
                    {
                        if (CheckIfBlankRow(iarr, k))
                        {
                            rowsToWrite = k + 1; // keep the first blank row, matching previous behaviour
                            break;
                        }
                    }
                }

                if (rowsToWrite > 0 && tillCol > 0)
                {
                    object[,] block = new object[rowsToWrite, tillCol];
                    for (int r = 0; r < rowsToWrite; r++)
                    {
                        for (int c = 0; c < tillCol; c++)
                        {
                            block[r, c] = iarr[r, c];
                        }
                    }

                    int startRow = ReportAppendRow;
                    WriteBlock(ReportWS, startRow, 1, block);
                    ReportAppendRow = startRow + rowsToWrite;
                }

                Range range = ReportWS.Range["A1", $"J{Math.Max(1, ReportAppendRow - 1)}"];
                FormatRange(range);

                ReportWB.Save();

                return ReportWS.Name;
            }
            catch (Exception ex)
            {
                Logging.Error("An error occurred: " + ex.Message);
                return "null";
            }
        }
        
        // this method Append data to Excel for ContinuityComponents report
        public string AppendContinuityComponentsandPartNumber(Dictionary<string, List<string>> matchedComponents)
        {
            try
            {
                SuspendExcelUpdates();
                int row = 2; // Starting from row 2, as row 1 contains headers ("Component Name" and "Part Number")
                var lines = new List<object[]>();
                var merges = new List<(int StartRow, int PartCount)>();
                int excelRow = row;

                foreach (var component in matchedComponents)
                {
                    int partCount = component.Value.Count;
                    if (partCount <= 0)
                        continue;

                    int startRow = excelRow;
                    for (int p = 0; p < partCount; p++)
                    {
                        object[] line = new object[2];
                        line[0] = p == 0 ? component.Key : "";
                        line[1] = component.Value[p];
                        lines.Add(line);
                        excelRow++;
                    }
                    merges.Add((startRow, partCount));
                }

                if (lines.Count > 0)
                {
                    object[,] block = new object[lines.Count, 2];
                    for (int r = 0; r < lines.Count; r++)
                    {
                        block[r, 0] = lines[r][0];
                        block[r, 1] = lines[r][1];
                    }
                    WriteBlock(ReportWS, row, 1, block);

                    for (int m = 0; m < merges.Count; m++)
                    {
                        var merge = merges[m];
                        Range mergeRange = ReportWS.Range[
                            ToA1(merge.StartRow, 1) + ":" + ToA1(merge.StartRow + merge.PartCount - 1, 1)];
                        mergeRange.Merge();
                        mergeRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        mergeRange.VerticalAlignment = XlVAlign.xlVAlignCenter;
                        ReleaseCom(mergeRange);
                    }
                }

                ReportWB.Save();
                Logging.Info($"Continuity Components: rows={lines.Count}");
                return ReportWS.Name;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Excel ##09: " + ex.Message);
                Logging.Error("An error occurred: " + ex.Message);
                return "null";
            }
            finally
            {
                DiscardReportWorkbook(save: true);
            }
        }

        private bool CheckIfBlankRow(object[,] iarr, int currentRow)
        {
            try
            {
                // Check if the current row and the next two rows are blank in columns 1 to 4
                return IsRowBlank(iarr, currentRow) &&
                       IsRowBlank(iarr, currentRow + 1) &&
                       IsRowBlank(iarr, currentRow + 2);
            }
            catch
            {
                return false;
            }
        }

        private bool IsRowBlank(object[,] iarr, int row)
        {
            // Check if all cells in columns 1 to 4 of the specified row are blank
            return row < iarr.GetLength(0) &&
                   string.IsNullOrEmpty(iarr[row, 0]?.ToString()) &&
                   string.IsNullOrEmpty(iarr[row, 1]?.ToString()) &&
                   string.IsNullOrEmpty(iarr[row, 2]?.ToString()) &&
                   string.IsNullOrEmpty(iarr[row, 3]?.ToString());
        }

        private void FormatRange(Range range)
        {
            range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            range.VerticalAlignment = XlVAlign.xlVAlignCenter;
            range.WrapText = true;
            range.Orientation = 0;
            range.AddIndent = false;
            range.IndentLevel = 0;
            range.ShrinkToFit = false;
            range.ReadingOrder = (int)Constants.xlContext;
            range.MergeCells = false;
        }
        #endregion
        // this method sort data an Excel sheet based on wire code for sheet,project wirelist, Component specifice breakdown reports
        public void SortWirelistWireNumber(int rowStart, int rowEnd)
        {
            Excel.Range sortRange = null;
            Excel.Range keyRange = null;
            try
            {
                string sRange = $"A{rowStart}:J{rowEnd}";
                string colFilter1 = $"E{rowStart}:E{rowEnd}";

                // Define the range to sort
                sortRange = ReportWS.Range[sRange];
                keyRange = ReportWS.Range[colFilter1];

                // Clear any previous sort fields
                ReportWS.Sort.SortFields.Clear();

                // Add the sorting field based on column E
                ReportWS.Sort.SortFields.Add(
                    Key: keyRange,
                    SortOn: XlSortOn.xlSortOnValues,
                    Order: XlSortOrder.xlAscending
                // DataOption: XlSortDataOption.xlSortNormal
                );

                // Apply sorting properties and execute the sort
                ReportWS.Sort.SetRange(sortRange);
                ReportWS.Sort.Header = XlYesNoGuess.xlGuess;
                ReportWS.Sort.MatchCase = false;
                ReportWS.Sort.Orientation = XlSortOrientation.xlSortColumns;
                ReportWS.Sort.SortMethod = XlSortMethod.xlPinYin;
                ReportWS.Sort.Apply();
                // Save the workbook
                ReportWB.Save();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Excel ##10: " + ex.Message);
                Logging.Error("An error occurred: " + ex.Message);
            }
            finally
            {
                DiscardReportWorkbook(save: true);
            }
        }
        /*public void SortWirelistWireNumber(int rowStart, int rowEnd) //commented on MAy 15th
        {
            Excel.Range sortRange = null;
            Excel.Range keyRange = null;
            try
            {
                ReportWS.Activate(); // Must activate before sorting
                string sRange = $"A{rowStart}:J{rowEnd}";
                string colFilter1 = $"E{rowStart}:E{rowEnd}";

                // Define the range to sort
                sortRange = ReportWS.Range[sRange];
                keyRange = ReportWS.Range[colFilter1];

                // Clear any previous sort fields
                ReportWS.Sort.SortFields.Clear();

                // Add the sorting field based on column E
                ReportWS.Sort.SortFields.Add(
                      keyRange,
                      XlSortOn.xlSortOnValues,
                      XlSortOrder.xlAscending,
                      Type.Missing,
                      XlSortDataOption.xlSortNormal
                );

                // Apply sorting properties and execute the sort
                ReportWS.Sort.SetRange(sortRange);
                ReportWS.Sort.Header = XlYesNoGuess.xlGuess;
                ReportWS.Sort.MatchCase = false;
                ReportWS.Sort.Orientation = XlSortOrientation.xlSortColumns;
                //ReportWS.Sort.Orientation = XlSortOrientation.xlSortRows;
                ReportWS.Sort.SortMethod = XlSortMethod.xlPinYin;

                ReportWS.Sort.Apply();
                // Save the workbook
                ReportWB.Save();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Excel ##10: "+ ex.Message);
                Logging.Error("An error occurred: " + ex.Message);
            }
        }*/

        // old method commented on MAy 9th 25
        /*  public int MaxRowinArray(object[,] iarr)
          {
              int K;
              int lowerBoundRow = iarr.GetLowerBound(0); // Rows
              int upperBoundRow = iarr.GetUpperBound(0); // Rows
              int lowerBoundCol = iarr.GetLowerBound(1); // Columns

              // Equivalent to 'On Error Resume Next' in VB
              for (K = lowerBoundRow; K <= upperBoundRow; K++)
              {
                  try
                  {
                      if ((iarr[K, 1] == null || iarr[K, 1].ToString() == "") &&
                          (iarr[K, 2] == null || iarr[K, 2].ToString() == "") &&
                          (iarr[K, 3] == null || iarr[K, 3].ToString() == "") &&
                          (iarr[K, 4] == null || iarr[K, 4].ToString() == "") &&

                          (iarr[K + 1, 1] == null || iarr[K + 1, 1].ToString() == "") &&
                          (iarr[K + 1, 2] == null || iarr[K + 1, 2].ToString() == "") &&
                          (iarr[K + 1, 3] == null || iarr[K + 1, 3].ToString() == "") &&
                          (iarr[K + 1, 4] == null || iarr[K + 1, 4].ToString() == "") &&

                          (iarr[K + 2, 1] == null || iarr[K + 2, 1].ToString() == "") &&
                          (iarr[K + 2, 2] == null || iarr[K + 2, 2].ToString() == "") &&
                          (iarr[K + 2, 3] == null || iarr[K + 2, 3].ToString() == "") &&
                          (iarr[K + 2, 4] == null || iarr[K + 2, 4].ToString() == ""))
                      {
                          break; // Exit the loop
                      }
                  }
                  catch (IndexOutOfRangeException ex)
                  {
                      MessageBox.Show("Excel ##11: "+ ex.Message);
                      // Ignore index out of range errors, similar to 'On Error Resume Next' in VB
                  }
              }

              return K - 1;
          }*/

        public int MaxRowinArray(object[,] iarr)
        {
            if (iarr == null)
                throw new ArgumentNullException(nameof(iarr));

            int lowerBoundRow = iarr.GetLowerBound(0); // usually 0 or 1
            int upperBoundRow = iarr.GetUpperBound(0); // max row index
            int lowerBoundCol = iarr.GetLowerBound(1); // usually 0 or 1
            int upperBoundCol = iarr.GetUpperBound(1);

            for (int K = lowerBoundRow; K <= upperBoundRow; K++)
            {
                bool isCurrentAndNext2RowsEmpty = true;

                // Check only if we have 3 rows ahead
                if (K + 2 <= upperBoundRow)
                {
                    for (int i = 0; i < 3; i++) // Check K, K+1, K+2
                    {
                        for (int j = 1; j <= 4; j++) // Columns 1 to 4
                        {
                            if (iarr[K + i, j] != null && iarr[K + i, j].ToString().Trim() != "")
                            {
                                isCurrentAndNext2RowsEmpty = false;
                                break;
                            }
                        }
                        if (!isCurrentAndNext2RowsEmpty)
                            break;
                    }

                    if (isCurrentAndNext2RowsEmpty)
                        return K - 1;
                }
            }

            // If we reached the end of the loop without breaking early, return the last row
            return upperBoundRow;
        }

        #region merging and Demerging

        public void SubLoomMergeCellsWithaValue(object iFrom, object iTo, object iVal1, object iVal2, object iNumOfShieldRemToMerge, object iCableClassification, Worksheet ws = null, object lengthValue = null)
        {
            ws ??= (Worksheet)ReportWB.ActiveSheet;
            CableListMerges.MergeMe(ws, iFrom, iTo, iVal1, iVal2, iNumOfShieldRemToMerge, iCableClassification, lengthValue);
        }

        public void SubLoomUNmergeCellsWithaValue(object iFrom, object iTo, object iVal1, object iVal2, object iCableClassification, Worksheet ws = null)
        {
            ws ??= (Worksheet)ReportWB.ActiveSheet;
            CableListMerges.UnmergeRow(ws, iFrom, iVal1, iVal2, iCableClassification);
        }
        #endregion

        private string PrepareCableListWorkbookFromTemplate(int numberOfSheetsRequired, bool useXlsx, string reportFilePath)
        {
            ExcelApp.DisplayAlerts = false;
            Worksheet keep = FindSheetByName(ReportWB, "CABLE LIST") ?? FindSheetByName(ReportWB, "CL-1");
            if (keep == null)
                throw new InvalidOperationException("Sheet 'CABLE LIST' does not exist in the template workbook.");

            ExcelSheetOps.KeepOnlyNamedSheet(ReportWB, keep);

            Worksheet first = (Worksheet)ReportWB.Sheets[1];
            if (first.Name != "CL-1")
                first.Name = "CL-1";

            first.Range["B47"].Value2 = DateTime.Now.Year;
            first.Range["B3:D4"].FormulaR1C1 = "";
            alignCellsXl(8, 46, first);

            // Copy in native .xlsx — .xls compatibility mode makes Sheet.Copy very slow.
            string xlsxWork = useXlsx
                ? reportFilePath
                : Path.Combine(Path.GetDirectoryName(reportFilePath) ?? "", Path.GetFileNameWithoutExtension(reportFilePath) + "_hal.work.xlsx");
            if (!useXlsx && File.Exists(xlsxWork))
                File.Delete(xlsxWork);
            ReportWB.SaveAs(xlsxWork, XlFileFormat.xlOpenXMLWorkbook);

            DuplicateTemplateSheets(numberOfSheetsRequired, "CL-");
            return xlsxWork;
        }

        private void DuplicateTemplateSheets(int needed, string namePrefix)
        {
            ExcelSheetOps.DuplicateTemplateSheets(ExcelApp, ReportWB, needed, namePrefix);
        }

        private void ApplyCl1HeaderFormulas(Worksheet ws)
        {
            CableListHeaderFormulas.ApplyCl1Formulas(ws);
        }

        private void CopyRangeToSheetGroup(Worksheet source, string address, int fromSheet, int toSheet)
        {
            ExcelSheetOps.CopyRangeToSheetGroup(ExcelApp, ReportWB, source, address, fromSheet, toSheet);
        }

        private void ApplyStrikethroughRows(Worksheet ws, List<int> excelRows)
        {
            CableListStrikethrough.Apply(ExcelApp, ws, excelRows);
        }

        public void GenerateHALReportFormat_CableList(object[,] iarr, string reportName, int numberOfSheetsRequired, int loomSheetRowReq, string reportType)
        {
            string templatePath = Environment.GetEnvironmentVariable("ELECTRE_CUSTOMIZE") + @"\system\CABLE_ReportFormat.xls";
            var sw = Stopwatch.StartNew();
            string workingCopy = null;
            string workXlsx = null;
            bool useXlsx = false;

            try
            {
                SuspendExcelUpdates();
                ExcelApp.DisplayAlerts = false;

                if (numberOfSheetsRequired < 1)
                    numberOfSheetsRequired = 1;

                string reportPath = Path.Combine(GlobalVar.ReportFolderGlobal, reportType);
                if (!Directory.Exists(reportPath))
                    Directory.CreateDirectory(reportPath);

                useXlsx = numberOfSheetsRequired > 250;
                string reportFilePath = Path.Combine(reportPath, reportName + (useXlsx ? ".xlsx" : ".xls"));
                XlFileFormat fileFormat = useXlsx ? XlFileFormat.xlOpenXMLWorkbook : XlFileFormat.xlExcel8;
                if (useXlsx)
                    Logging.Info($"Cable list '{reportName}' has {numberOfSheetsRequired} sheets; saving as .xlsx because .xls is limited to 255 sheets.");

                workingCopy = Path.Combine(reportPath, reportName + "_hal.tmp.xls");
                if (File.Exists(workingCopy))
                    File.Delete(workingCopy);
                File.Copy(templatePath, workingCopy, true);

                ReportWB = ExcelApp.Workbooks.Open(workingCopy);
                workXlsx = PrepareCableListWorkbookFromTemplate(numberOfSheetsRequired, useXlsx, reportFilePath);
                long copyMs = sw.ElapsedMilliseconds;
                Logging.Info($"Cable list '{reportName}': copied {numberOfSheetsRequired} sheets in {copyMs} ms");

                int Q = 1;
                int maxQ = iarr.GetLength(0) - 1;
                LoomInitialRow = 8;
                LoomInitialCol = 2;
                double dateColWidth = 0;
                Worksheet headerTemplateSheet = null;

                for (int R = 1; R <= numberOfSheetsRequired; R++)
                {
                    Worksheet ws1 = (Worksheet)ReportWB.Sheets[R];

                    int rowsThisSheet = 0;
                    if (Q <= maxQ && loomSheetRowReq > 0)
                        rowsThisSheet = Math.Min(loomSheetRowReq, maxQ - Q + 1);

                    if (rowsThisSheet > 0)
                    {
                        object[,] block = CableListDataBlock.Build(iarr, Q, rowsThisSheet);

                        int startExcelRow = LoomInitialRow + 1;
                        WriteBlock(ws1, startExcelRow, LoomInitialCol + 1, block);

                        var strikeRows = new List<int>();
                        var mergedMe = new HashSet<string>();
                        var meOps = new List<HalMeMerge>();
                        for (int r = 0; r < rowsThisSheet; r++)
                        {
                            int q = Q + r;
                            int excelRow = startExcelRow + r;

                            if (string.Equals(iarr[q, 19]?.ToString(), "D", StringComparison.Ordinal))
                                strikeRows.Add(excelRow);

                            if (!string.IsNullOrEmpty(iarr[q, 12] as string))
                            {
                                string mergeKind = iarr[q, 11]?.ToString();
                                if (mergeKind == "ME")
                                {
                                    string meKey = Convert.ToString(iarr[q, 12]) + ":" + Convert.ToString(iarr[q, 13]) + ":"
                                        + Convert.ToString(iarr[q, 16]) + ":" + Convert.ToString(iarr[q, 18]);
                                    if (!mergedMe.Add(meKey))
                                        continue;
                                    meOps.Add(CableListMerges.BuildMe(iarr[q, 12], iarr[q, 13], iarr[q, 14], iarr[q, 16], iarr[q, 18]));
                                }
                                else if (mergeKind == "UN")
                                {
                                    SubLoomUNmergeCellsWithaValue(iarr[q, 12], iarr[q, 12], iarr[q, 14], iarr[q, 15], iarr[q, 18], ws1);
                                }
                                else if (mergeKind == "LAST")
                                {
                                    ws1.Range["B" + excelRow].Value2 = iarr[q, 14];
                                }
                            }
                        }

                        CableListMerges.ApplyMeBatch(ws1, meOps);
                        ApplyStrikethroughRows(ws1, strikeRows);
                        Q += rowsThisSheet;
                    }

                    object[,] sheetLabel = new object[2, 1];
                    sheetLabel[0, 0] = $"SHEET {R} OF {numberOfSheetsRequired} SHEETS";
                    sheetLabel[1, 0] = sheetLabel[0, 0];
                    ws1.Range["L51:L52"].Value2 = sheetLabel;
                    if (R == 1)
                    {
                        if (reportName.Length > 1)
                        {
                            ws1.Range["F51:F52"].Value = reportName.Substring(0, reportName.Length - 1);
                            ws1.Range["K52"].Value = reportName.Substring(reportName.Length - 1);
                        }
                        else
                        {
                            ws1.Range["F51:F52"].Value = reportName;
                            ws1.Range["K52"].Value = "";
                        }
                    }
                    else if (R == 2)
                    {
                        ApplyCl1HeaderFormulas(ws1);
                        ((Range)ws1.Columns["D"]).AutoFit();
                        dateColWidth = Convert.ToDouble(((Range)ws1.Columns["D"]).ColumnWidth);
                        headerTemplateSheet = ws1;
                    }
                }

                if (numberOfSheetsRequired > 2 && headerTemplateSheet != null)
                {
                    CopyRangeToSheetGroup(headerTemplateSheet, "C48:K52", 3, numberOfSheetsRequired);
                    if (dateColWidth > 0)
                    {
                        SelectSheetGroup(3, numberOfSheetsRequired);
                        ((Range)((Worksheet)ExcelApp.ActiveSheet).Columns["D"]).ColumnWidth = dateColWidth;
                        UngroupSheets();
                    }
                }

                SelectSheetGroup(1, numberOfSheetsRequired);
                alignCellsXl(8, 46, (Worksheet)ExcelApp.ActiveSheet);
                UngroupSheets();
                ClearCutCopyMode();

                long fillMs = sw.ElapsedMilliseconds - copyMs;
                long saveStart = sw.ElapsedMilliseconds;
                if (useXlsx)
                    ReportWB.Save();
                else
                    ReportWB.SaveAs(reportFilePath, fileFormat);
                long saveMs = sw.ElapsedMilliseconds - saveStart;
                Logging.Info($"Cable list '{reportName}': copy={copyMs}ms fill={fillMs}ms save={saveMs}ms total={sw.ElapsedMilliseconds}ms sheets={numberOfSheetsRequired} -> {reportFilePath}");
            }
            catch (COMException ex)
            {
                MessageBox.Show("Excel ##12: "+ ex.Message);
                Logging.Error("Error interacting with Excel: " + ex.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Excel ##13: "+ ex.Message);
                Logging.Error("An error occurred: " + ex.Message);
            }
            finally
            {
                if (TempWB != null)
                {
                    try { TempWB.Close(false); } catch { }
                    Marshal.ReleaseComObject(TempWB);
                    TempWB = null;
                }
                if (ReportWB != null)
                {
                    ReportWB.Close(false);
                    Marshal.ReleaseComObject(ReportWB);
                    ReportWB = null;
                }
                if (workingCopy != null && File.Exists(workingCopy))
                {
                    try { File.Delete(workingCopy); } catch { }
                }
                if (!useXlsx && workXlsx != null && File.Exists(workXlsx))
                {
                    try { File.Delete(workXlsx); } catch { }
                }
            }
        }
        
        public void alignCellsXl(int fromCell, int toCell, Worksheet ws = null)
        {
            ws ??= (Worksheet)ReportWB.ActiveSheet;
            Range rangeB = ws.Range["B" + fromCell, "O" + toCell];
            rangeB.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            rangeB.VerticalAlignment = XlVAlign.xlVAlignCenter;
        }
        public int GetNumberOfSheetsRequired(int iTotalLines, int iNoLinesCanBePrinted)
        {
            double X = (double)iTotalLines / iNoLinesCanBePrinted;

            if (X == Math.Round(X, 0))
            {
                return (int)X;
            }
            else
            {
                return (int)Math.Round(X + 0.5, 0);
            }
        }

        // this method generates Continuity List report format with HAL template
        private string PrepareContinuityWorkbookFromTemplate(int numberOfSheetsRequired, bool useXlsx, string reportFilePath)
        {
            ExcelApp.DisplayAlerts = false;
            Worksheet keep = FindSheetByName(ReportWB, "continuity", ignoreCase: true);
            if (keep == null)
                throw new InvalidOperationException("Sheet 'continuity' does not exist in the template workbook.");

            ExcelSheetOps.KeepOnlyNamedSheet(ReportWB, keep);

            Worksheet first = (Worksheet)ReportWB.Sheets[1];
            if (first.Name != "CWOB-1")
                first.Name = "CWOB-1";

            alignCellsXl(6, 53, first);

            string xlsxWork = useXlsx
                ? reportFilePath
                : Path.Combine(Path.GetDirectoryName(reportFilePath) ?? "", Path.GetFileNameWithoutExtension(reportFilePath) + "_cwob.work.xlsx");
            if (!useXlsx && File.Exists(xlsxWork))
                File.Delete(xlsxWork);
            ReportWB.SaveAs(xlsxWork, XlFileFormat.xlOpenXMLWorkbook);

            DuplicateTemplateSheets(numberOfSheetsRequired, "CWOB-");
            return xlsxWork;
        }

        private void ApplyCwob1HeaderFormulas(Worksheet ws)
        {
            ContinuityHeaderFormulas.ApplyCwob1Formulas(ws);
        }

        private void WriteContinuityDataBlock(Worksheet ws, int startRow, object[,] arrFTcwob, int qStart, int rows)
        {
            ContinuityDataBlock.Write(ws, startRow, arrFTcwob, qStart, rows);
        }

        public void GenerateHALReportFormat_ContinuityList(object[,] arrFTcwob, string reportType, string reportName)
        {
            int numRowInReportSheet = 47;
            int numberOfSheetsRequired = 0;

            int arrFTcwobCount = 0;

            for (int i = 0; i < arrFTcwob.GetLength(0); i++)
            {
                if (!string.IsNullOrEmpty((string)arrFTcwob[i, 0]) || !string.IsNullOrEmpty((string)arrFTcwob[i, 1]) || !string.IsNullOrEmpty((string)arrFTcwob[i, 2]) || !string.IsNullOrEmpty((string)arrFTcwob[i, 3]) || !string.IsNullOrEmpty((string)arrFTcwob[i, 4]) || !string.IsNullOrEmpty((string)arrFTcwob[i, 5]) || !string.IsNullOrEmpty((string)arrFTcwob[i, 6]))//removing the empty and null values
                {
                    arrFTcwobCount++;
                }
            }

            numberOfSheetsRequired = arrFTcwobCount / (numRowInReportSheet + 1) + 1;

            string templatePath = Environment.GetEnvironmentVariable("ELECTRE_CUSTOMIZE") + @"\system\CONTINUITY_ReportFormat.xls";

            if (string.IsNullOrEmpty(templatePath))
            {
                MessageBox.Show("Continuity Report template Not Found on the System Folder", "Template Not Found", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var sw = Stopwatch.StartNew();
            string workingCopy = null;
            string workXlsx = null;
            bool useXlsx = false;

            try
            {
                SuspendExcelUpdates();
                ExcelApp.DisplayAlerts = false;

                if (numberOfSheetsRequired < 1)
                    numberOfSheetsRequired = 1;

                string reportPath = Path.Combine(GlobalVar.ReportFolderGlobal, reportType);
                if (!Directory.Exists(reportPath))
                    Directory.CreateDirectory(reportPath);

                useXlsx = numberOfSheetsRequired > 250;
                string reportFilePath = Path.Combine(reportPath, reportName + (useXlsx ? ".xlsx" : ".xls"));
                XlFileFormat fileFormat = useXlsx ? XlFileFormat.xlOpenXMLWorkbook : XlFileFormat.xlExcel8;
                if (useXlsx)
                    Logging.Info($"Continuity '{reportName}' has {numberOfSheetsRequired} sheets; saving as .xlsx because .xls is limited to 255 sheets.");

                workingCopy = Path.Combine(reportPath, reportName + "_cwob.tmp.xls");
                if (File.Exists(workingCopy))
                    File.Delete(workingCopy);
                File.Copy(templatePath, workingCopy, true);

                ReportWB = ExcelApp.Workbooks.Open(workingCopy);
                workXlsx = PrepareContinuityWorkbookFromTemplate(numberOfSheetsRequired, useXlsx, reportFilePath);
                long copyMs = sw.ElapsedMilliseconds;
                Logging.Info($"Continuity '{reportName}': copied {numberOfSheetsRequired} sheets in {copyMs} ms");

                int Q = 0;
                int maxQ = arrFTcwob.GetLength(0) - 1;
                int StartRow = 6;
                int rowsPerSheet = numRowInReportSheet + 1; // rows 6..53 inclusive
                Worksheet headerTemplateSheet = null;

                for (int R = 1; R <= numberOfSheetsRequired; R++)
                {
                    Worksheet ws1 = (Worksheet)ReportWB.Sheets[R];
                    ws1.Range["L4"].Value2 = $"{R} OF {numberOfSheetsRequired}";

                    int rowsThisSheet = 0;
                    if (Q <= maxQ)
                        rowsThisSheet = Math.Min(rowsPerSheet, maxQ - Q + 1);

                    if (rowsThisSheet > 0)
                    {
                        WriteContinuityDataBlock(ws1, StartRow, arrFTcwob, Q, rowsThisSheet);
                        Q += rowsThisSheet;
                    }

                    if (R == 2)
                    {
                        ApplyCwob1HeaderFormulas(ws1);
                        headerTemplateSheet = ws1;
                    }
                }

                if (numberOfSheetsRequired > 2 && headerTemplateSheet != null)
                {
                    CopyRangeToSheetGroup(headerTemplateSheet, "F4:I4", 3, numberOfSheetsRequired);
                    CopyRangeToSheetGroup(headerTemplateSheet, "C56:L57", 3, numberOfSheetsRequired);
                }

                ClearCutCopyMode();

                long fillMs = sw.ElapsedMilliseconds - copyMs;
                long saveStart = sw.ElapsedMilliseconds;
                if (useXlsx)
                    ReportWB.Save();
                else
                    ReportWB.SaveAs(reportFilePath, fileFormat);
                long saveMs = sw.ElapsedMilliseconds - saveStart;
                Logging.Info($"Continuity '{reportName}': copy={copyMs}ms fill={fillMs}ms save={saveMs}ms total={sw.ElapsedMilliseconds}ms sheets={numberOfSheetsRequired} -> {reportFilePath}");
            }
            catch (COMException ex)
            {
                MessageBox.Show("Excel ##14: "+ ex.Message);
                Logging.Error("Error interacting with Excel: " + ex.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Excel ##15: "+ ex.Message);
                Logging.Error("An error occurred: " + ex.Message);
            }
            finally
            {
                if (TempWB != null)
                {
                    try { TempWB.Close(false); } catch { }
                    Marshal.ReleaseComObject(TempWB);
                    TempWB = null;
                }
                if (ReportWB != null)
                {
                    ReportWB.Close(false);
                    Marshal.ReleaseComObject(ReportWB);
                    ReportWB = null;
                }
                if (workingCopy != null && File.Exists(workingCopy))
                {
                    try { File.Delete(workingCopy); } catch { }
                }
                if (!useXlsx && workXlsx != null && File.Exists(workXlsx))
                {
                    try { File.Delete(workXlsx); } catch { }
                }
            }
        }
        // after generating Excel reports this method releseComObjactes
        public void releaseExcel()
        {
            if (ExcelApp == null)
                return;

            ResumeExcelUpdates();
            ExcelApp.Quit();
            Marshal.ReleaseComObject(ExcelApp);
            ExcelApp = null;
        }

      // this method create header for sheet,project wirelist,component specifice breakdown reports 
        public bool CreateWirelistReportHeader(string workbookPath)
        {
            ReportAppendRow = 1;

            try
            {
                SuspendExcelUpdates();
                if (!IsWorkbookOpen(workbookPath))
                    ReportWB = ExcelApp.Workbooks.Open(workbookPath);

                ReportWS = ReportWB.ActiveSheet as Worksheet;

                if (ReportWS == null)
                {
                    Logging.Error("Failed to access the active sheet.");
                    return false;
                }

                object[,] header = new object[1, 10];
                header[0, 0] = "FROM CONN";
                header[0, 1] = "FROM PIN";
                header[0, 2] = "TO CONN";
                header[0, 3] = "TO PIN";
                header[0, 4] = "WIRE CODE";
                header[0, 5] = "WIRE TYPE";
                header[0, 6] = "LENGTH";
                header[0, 7] = "RD";
                header[0, 8] = "LD";
                header[0, 9] = " NERD";
                WriteBlock(ReportWS, ReportAppendRow, 1, header);

                Range headerRange = ReportWS.Range["A1", "J1"];
                headerRange.Font.Bold = true;
                ReleaseCom(headerRange);

                ((Range)ReportWS.Columns["A"]).ColumnWidth = 14.78;
                ((Range)ReportWS.Columns["B"]).ColumnWidth = 14.78;
                ((Range)ReportWS.Columns["C"]).ColumnWidth = 14;
                ((Range)ReportWS.Columns["D"]).ColumnWidth = 16.67;
                ((Range)ReportWS.Columns["E"]).ColumnWidth = 21.33;
                ((Range)ReportWS.Columns["F"]).ColumnWidth = 14.11;
                ((Range)ReportWS.Columns["G"]).ColumnWidth = 10;
                ((Range)ReportWS.Columns["H"]).ColumnWidth = 20;
                ((Range)ReportWS.Columns["I"]).ColumnWidth = 15;
                ((Range)ReportWS.Columns["J"]).ColumnWidth = 10;

                ReportAppendRow++;

                Logging.Info($"CreateBreakdownReportHeader created for {workbookPath}");
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Excel ##17: " + ex.Message);
                Logging.Error("Failed to create CreateBreakdownReportHeader: " + ex.Message);
                return false;
            }
        }

        private List<ElectreObject> GetSCB_TCB_Source(List<ElectreObject> panelCollection)
        {
            return panelCollection
                .Where(e => e.ComponentType.Contains("SCB", StringComparison.OrdinalIgnoreCase) ||
                            e.ComponentType.Contains("TCB", StringComparison.OrdinalIgnoreCase))
               .OrderBy(e => e.ComponentType.Contains("SCB", StringComparison.OrdinalIgnoreCase) ? 0 : 1)  // SCB first
                .ThenBy(e => e.Panel)
                  .ThenBy(e => e.ConnectorName)
                 .ToList();
        }

        public string AppendToExcelPowerOn(List<ElectreObject> elecCollection)
        {
            List<string> parameterPINs = new List<string>();
            var sw = Stopwatch.StartNew();
            try
            {
                SuspendExcelUpdates();
                _powerOnBlocks = new List<PowerOnExcelBlock>();

                PowerOnReport.AllIndex = ElectreTraceIndex.Build(modMain.ElecCollection_All);
                var sourceComponents = GetSCB_TCB_Source(elecCollection);
                if (sourceComponents.Count == 0)
                {
                    MessageBox.Show("Circuit breakers not found in the Panel", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Logging.Info("Circuit breakers not found in the Panel");
                    return "null";
                }

                foreach (var source in sourceComponents)
                {
                    var visitedConnections = new HashSet<string>();

                    //Get the Negative/Ground pins for the source component
                    List<string> negPins = PowerOnReport.FindGroundPin(elecCollection, source);

                    // Tracing the Destination component from Source
                    TraceAndLogPath(source, source.WireNumber, source.ConnectorName, source.PinNumber, source.SubNet, visitedConnections, negPins, ref parameterPINs);
                }

                FlushPowerOnBuffer();

                if (parameterPINs.Count >= 2)
                {
                    // Join the values as strings with a hyphen
                    string result = string.Join("-", parameterPINs.Select(pin => pin.ToString()));

                    // Set the cell format to Text to prevent Excel from interpreting it as a date
                    Range paramCell = ReportWS.Range["E" + ReportAppendRow];
                    paramCell.NumberFormat = "@";  // Set to Text format
                    paramCell.Value2 = result;
                    paramCell.Font.Bold = true;
                    ReleaseCom(paramCell);

                    ReportAppendRow += 2;
                    parameterPINs.Clear();
                }
                ((Range)ReportWS.Columns[1]).AutoFit();
                ((Range)ReportWS.Columns[5]).AutoFit();
                ReportWB.Save();
                Logging.Info($"Power On '{ReportWS.Name}': total={sw.ElapsedMilliseconds}ms -> {ReportWB.FullName}");
                return ReportWS.Name;
            }
            catch (Exception ex)
            {
                Logging.Error("AppendToExcelPowerOn: Error while appending PowerOn data: " + ex.Message);
                return "null";
            }
            finally
            {
                _powerOnBlocks = null;
            }
        }

        // Helper Method: Recursive path tracing
        private void TraceAndLogPath(ElectreObject source, string wireNumber, string connectorName, string pinNumber, string subNet, HashSet<string> visited, List<string> negPins,ref List<string> parameterPINs)
        {
            try
            {
                visited.Add($"{connectorName},{pinNumber}");

                var wirePeers = PowerOnReport.AllIndex != null
                    ? PowerOnReport.AllIndex.ConnectedOnWire(wireNumber, subNet)
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
                    LogToExcel(source, connectorName, pinNumber, negPins, ref parameterPINs);
                    return;
                }

                foreach (var connObj in connectedObjects)
                {
                    if (visited.Contains($"{connObj.ConnectorName},{connObj.PinNumber}"))
                        continue;

                    switch (connObj.ComponentType)
                    {
                        case "EQU":
                            var nextObjEQU = PowerOnReport.TracePinofEQUConnector(connObj);
                            if (nextObjEQU == null)
                            {
                                LogToExcel(source, connObj.ConnectorName, connObj.PinNumber, negPins, ref parameterPINs);
                                return;
                            }
                            if (nextObjEQU != null)
                                TraceAndLogPath(source, nextObjEQU.WireNumber, nextObjEQU.ConnectorName, nextObjEQU.PinNumber, nextObjEQU.SubNet, visited, negPins, ref parameterPINs);
                            break;
                        case "DIS":
                            var nextObj = PowerOnReport.TracePinofBreakConnector(connObj);
                            if (nextObj == null)
                            {
                                LogToExcel(source, connObj.ConnectorName, connObj.PinNumber, negPins, ref parameterPINs);
                                return;
                            }
                            if (nextObj != null)
                                TraceAndLogPath(source, nextObj.WireNumber, nextObj.ConnectorName, nextObj.PinNumber, nextObj.SubNet, visited, negPins, ref parameterPINs);
                            break;

                        case "TBK":
                            var jmList = PowerOnReport.TracePinOfJM(connObj);
                            if (jmList.Count == 0)
                            {
                                LogToExcel(source, connObj.ConnectorName, connObj.PinNumber, negPins, ref parameterPINs);
                                return;
                            }
                            foreach (var jm in jmList)
                            {
                                TraceAndLogPath(source, jm.WireNumber, jm.ConnectorName, jm.PinNumber, jm.SubNet, visited, negPins, ref parameterPINs);
                            }
                            break;

                        case "SPL":
                            var splList = PowerOnReport.TracePinOfSPL(connObj);
                            if (splList.Count == 0)
                            {
                                LogToExcel(source, connObj.ConnectorName, connObj.PinNumber, negPins, ref parameterPINs);
                                return;
                            }
                            foreach (var spl in splList)
                            {
                                TraceAndLogPath(source, spl.WireNumber, spl.ConnectorName, spl.PinNumber, spl.SubNet, visited, negPins, ref parameterPINs);
                            }
                            break;

                        case "TER":
                            var terList = PowerOnReport.TracePinOfTER(connObj);
                            if (terList.Count == 0)
                            {
                                LogToExcel(source, connObj.ConnectorName, connObj.PinNumber, negPins, ref parameterPINs);
                                return;
                            }
                            foreach (var ter in terList)
                            {
                                TraceAndLogPath(source, ter.WireNumber, ter.ConnectorName, ter.PinNumber, ter.SubNet, visited, negPins, ref parameterPINs);
                            }
                            break;

                        default:
                            LogToExcel(source, connObj.ConnectorName, connObj.PinNumber, negPins, ref parameterPINs);
                            //TraceAndLogPath(source, connObj.WireNumber, connObj.ConnectorName, connObj.PinNumber, connObj.SubNet, visited, negPins, ref parameterPINs);
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                Logging.Error($"TraceAndLogPath Error: Source={source?.ConnectorName},{source?.PinNumber} and Dest={connectorName},{pinNumber} - {ex.Message}");
            }
        }

        // Helper Method: Logging to Excel
        private void LogToExcel(ElectreObject source, string destConnector, string destPin, List<string> negPins, ref List<string> parameterPINs)
        {
            try
            {
                int startRow = ReportAppendRow;
                var rows = new List<object[]>(4);

                object[] first = new object[5];
                first[0] = source.Panel;
                first[1] = source.ConnectorName;
                first[2] = destConnector;
                first[3] = $"{destPin}(+)";
                first[4] = source.Voltage;
                rows.Add(first);

                int i = 1;

                if (negPins.Count == 0)
                {
                    object[] groundRow = new object[5];
                    groundRow[1] = "";
                    groundRow[2] = "wrt STR";
                    groundRow[3] = "-";
                    rows.Add(groundRow);
                    i++;
                }
                else
                {
                    foreach (var negPin in negPins)
                    {
                        if (negPin.Contains(","))
                        {
                            string groundConnector = negPin.Split(',')[0];
                            string groundPin = negPin.Split(',')[1];
                            object[] groundRow = new object[5];
                            groundRow[1] = "";
                            groundRow[2] = groundConnector;
                            groundRow[3] = $"{groundPin}(-)";
                            rows.Add(groundRow);
                            i++;
                        }
                    }
                }

                int totalRows = i;
                object[,] block = new object[totalRows, 5];
                for (int r = 0; r < totalRows; r++)
                {
                    object[] row = rows[r];
                    for (int c = 0; c < 5; c++)
                        block[r, c] = row[c];
                }

                bool isScb = string.Equals(source.ComponentType, "SCB", StringComparison.OrdinalIgnoreCase);
                bool isTcb = string.Equals(source.ComponentType, "TCB", StringComparison.OrdinalIgnoreCase);
                _powerOnBlocks.Add(new PowerOnExcelBlock
                {
                    StartRow = startRow,
                    TotalRows = totalRows,
                    Rows = block,
                    MergePanelScb = isScb,
                    MergePanelTcb = isTcb
                });

                ReportAppendRow += totalRows + 1;

                if (isTcb)
                {
                    // adds if component type is TCB only
                    parameterPINs.Add(destPin);
                }
            }
            catch (Exception ex)
            {
                Logging.Error($"LogToExcel Error: Connector={destConnector}, Pin={destPin} - {ex.Message}");
            }
        }

        private void FlushPowerOnBuffer()
        {
            if (ReportWS == null || _powerOnBlocks == null || _powerOnBlocks.Count == 0)
                return;

            // Same order as the original per-call writes: values, then merges, then AutoFit.
            // TCB panel merge overlaps the next block's first row, so blocks cannot be merged after one sheet-wide dump.
            for (int b = 0; b < _powerOnBlocks.Count; b++)
            {
                PowerOnExcelBlock block = _powerOnBlocks[b];
                WriteBlock(ReportWS, block.StartRow, 1, block.Rows);

                int lastDataRow = block.StartRow + block.TotalRows - 1;
                Range mergeE = ReportWS.Range[ToA1(block.StartRow, 5) + ":" + ToA1(lastDataRow, 5)];
                mergeE.Merge();
                ReleaseCom(mergeE);

                Range mergeB = ReportWS.Range[ToA1(block.StartRow, 2) + ":" + ToA1(lastDataRow, 2)];
                mergeB.Merge();
                ReleaseCom(mergeB);

                if (block.MergePanelScb)
                {
                    Range mergeA = ReportWS.Range[ToA1(block.StartRow, 1) + ":" + ToA1(lastDataRow, 1)];
                    mergeA.Merge();
                    ReleaseCom(mergeA);
                }
                else if (block.MergePanelTcb)
                {
                    Range mergeA = ReportWS.Range[ToA1(block.StartRow, 1) + ":" + ToA1(block.StartRow + block.TotalRows + 1, 1)];
                    mergeA.Merge();
                    ReleaseCom(mergeA);
                }

                ((Range)ReportWS.Rows[block.StartRow]).AutoFit();
                ((Range)ReportWS.Rows[lastDataRow]).AutoFit();
            }
        }

        #region AppendToExcelPowerOn old methods
        /*public string AppendToExcelPowerOn(List<ElectreObject> elecCollection)
        {
            try
            {
                // Step 1: Get only SCBs & TCBs as source components
                var sourceComponents = GetSCB_TCB_Source(elecCollection);
                if (sourceComponents.Count == 0)
                {
                    MessageBox.Show($"Circuit breakers not found in the Panel", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Logging.Info("Circuit breakers not found in the Panel");
                }
                string ConnectorNameandPinTER = string.Empty;
                string breakConnectorName = string.Empty;

                List<string> parameterPINs = new List<string>();

                // Step 2: Loop through each SCB/TCB as a starting point
                foreach (var source in sourceComponents)
                {
                    string currentWireNumber = source.WireNumber;
                    string currentConnectorName = source.ConnectorName;
                    string currentPinNumber = source.PinNumber;
                   // string tag = source.Core_Part_Number;
                    string subNet = source.SubNet;

                    string groundConnectorName = string.Empty;
                    string groundPinNumber = string.Empty;

                    // Find the most relevant GND pin
                    string negPin = PowerOnReport.FindGroundPin(elecCollection, source);

                    if (!string.IsNullOrEmpty(negPin))
                    {
                        if (negPin.Contains(","))
                        {
                            string[] negPinParts = negPin.Split(',');
                            groundConnectorName = negPinParts[0];
                            groundPinNumber = negPinParts[1];
                        }
                    }

                    //  List<string> visited = new List<string>();
                    List<string> visitedConnections = new List<string>();

                    // Step 3: Trace the continuity path
                    while (!string.IsNullOrEmpty(currentWireNumber))
                    {
                        // visited.Add(visitKey);
                        visitedConnections.Add($"{currentConnectorName},{currentPinNumber}");

                        // Find all objects that share the same wire number and subnet
                        var connectedObjects = modMain.ElecCollection_All
                                                      .Where(w => w.WireNumber + "," + w.SubNet == currentWireNumber + "," + subNet &&
                                                      !visitedConnections.Contains($"{w.ConnectorName},{w.PinNumber}") && !string.IsNullOrEmpty(w.ConnectorName))
                                                      .ToList();

                        if (connectedObjects.Count == 0)
                            break; // Stop if no further connections or loop detected

                        bool foundNext = false;

                        foreach (var connObj in connectedObjects)
                        {
                            if (connObj.ConnectorName == currentConnectorName && connObj.PinNumber == currentPinNumber)
                                continue; // Skip self

                            // Update last known connector details
                            currentConnectorName = connObj.ConnectorName;
                            currentPinNumber = connObj.PinNumber;

                            
                            visitedConnections.Add($"{currentConnectorName},{currentPinNumber}");

                            switch (connObj.ComponentType)
                            {
                                case "EQU":
                                    currentConnectorName = $"{connObj.ConnectorName}";
                                    var ConnectedObjsEQU = PowerOnReport.TracePinofBreakConnector(connObj);
                                    if (ConnectedObjsEQU != null)
                                    {
                                        currentConnectorName = ConnectedObjsEQU.ConnectorName;
                                    }
                                    break;
                                case "DIS":
                                    var ConnectedObjsDIS = PowerOnReport.TracePinofBreakConnector(connObj);
                                    if (ConnectedObjsDIS != null)
                                    {
                                        currentConnectorName = ConnectedObjsDIS.ConnectorName;
                                    }                                    
                                    break;
                                case "TBK":
                                    var ConnectedObjsTBK = PowerOnReport.TracePinOfJM(connObj);
                                    foreach (var obj in ConnectedObjsTBK)
                                    {
                                        currentConnectorName = obj.ConnectorName;
                                        currentPinNumber = obj.PinNumber;
                                        connectedObjects.Add(obj);
                                    }
                                    break;
                                case "SPL":
                                    var ConnectedObjsSPL = PowerOnReport.TracePinOfSPL(connObj);
                                    foreach (var obj in ConnectedObjsSPL)
                                    {
                                        currentConnectorName = obj.ConnectorName;
                                        currentPinNumber = obj.PinNumber;
                                        connectedObjects.Add(obj);
                                    }
                                    break;
                                case "TER":
                                    var ConnectedObjsTER = PowerOnReport.TracePinOfTER(connObj);
                                    foreach(var obj in ConnectedObjsTER)
                                    {
                                        currentConnectorName = obj.ConnectorName;
                                        currentPinNumber = obj.PinNumber;
                                        connectedObjects.Add(obj);
                                    }
                                    break;
                                case "REL":
                                case "SWT":
                                case "DD":
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
                                default:
                                    currentConnectorName = connObj.ConnectorName;
                                    break;//pass the for the empty value
                            }
                            if (ConnectorNameandPinTER == "null")
                            {
                                break;
                            }
                            if (breakConnectorName == "null")
                            {
                                break;
                            }

                            // Step 4: Find the next wire connection based on the new connector
                            var nextWireConnection = modMain.ElecCollection_All
                                .FirstOrDefault(w => string.Equals(w.ConnectorName, currentConnectorName, StringComparison.OrdinalIgnoreCase) && w.PinNumber == currentPinNumber);

                            if (nextWireConnection != null)
                            {
                                // Move to the next wire
                                currentWireNumber = nextWireConnection.WireNumber;
                                currentConnectorName = nextWireConnection.ConnectorName;
                                currentPinNumber = nextWireConnection.PinNumber;
                               // tag = nextWireConnection.Core_Part_Number;
                                subNet = nextWireConnection.SubNet;

                                foundNext = true;
                                break; // Move to next wire
                            }
                        }

                        if (!foundNext)
                            break; // Stop if no further valid connections
                    }
                    ReportWS.Cells[ReportAppendRow, 1].Value = source.Panel;
                    ReportWS.Cells[ReportAppendRow, 2].Value = source.ConnectorName; // SCB/TCB Source
                    ReportWS.Cells[ReportAppendRow, 3].Value = currentConnectorName; // Final Destination Unit                
                    ReportWS.Cells[ReportAppendRow, 4].Value = $"{currentPinNumber}(+)"; // positive destination pin

                    if (string.IsNullOrEmpty(negPin))
                    {
                        ReportWS.Cells[ReportAppendRow + 1, 2].Value = "";
                        ReportWS.Cells[ReportAppendRow + 1, 3].Value = "wrt STR";
                        ReportWS.Cells[ReportAppendRow + 1, 4].Value = "-"; 
                    }
                    else if (!string.IsNullOrEmpty(negPin))
                    {
                        ReportWS.Cells[ReportAppendRow + 1, 2].Value = "";
                        ReportWS.Cells[ReportAppendRow + 1, 3].Value = $"{groundConnectorName}"; // Final Destination Unit for ground
                        ReportWS.Cells[ReportAppendRow + 1, 4].Value = $"{groundPinNumber}(-)";  // negative destination pin
                    }

                    ReportWS.Cells[ReportAppendRow, 5].Value = source.Voltage; // Parameter Column

                    ReportWS.Range[ReportWS.Cells[ReportAppendRow, 5], ReportWS.Cells[ReportAppendRow + 1, 5]].Merge();  // Merge the rows in column 5


                    // Merge the rows in column 1
                    if (string.Equals(source.ComponentType, "SCB", StringComparison.OrdinalIgnoreCase))
                    {
                        ReportWS.Range[ReportWS.Cells[ReportAppendRow, 1], ReportWS.Cells[ReportAppendRow + 1, 1]].Merge();
                    }
                    else if(string.Equals(source.ComponentType, "TCB", StringComparison.OrdinalIgnoreCase))
                    {
                        ReportWS.Range[ReportWS.Cells[ReportAppendRow, 1], ReportWS.Cells[ReportAppendRow + 3, 1]].Merge();
                    }

                    // Merge the rows in column 2
                    ReportWS.Range[ReportWS.Cells[ReportAppendRow, 2], ReportWS.Cells[ReportAppendRow + 1, 2]].Merge();

                    // ReportWS.Rows.AutoFit();  // Auto-size column width
                    ReportWS.Rows[ReportAppendRow].AutoFit();
                    ReportWS.Rows[ReportAppendRow + 1].AutoFit();

                    // Auto-fit for 4th Column "Parameter"
                    // ReportWS.Columns[4].AutoFit();

                    ReportAppendRow += 3;

                    if (string.Equals(source.ComponentType, "TCB", StringComparison.OrdinalIgnoreCase))
                    {
                        // adds if component type is TCB only
                        parameterPINs.Add(currentPinNumber);
                    }

                    if (parameterPINs.Count == 3)
                    {
                        // Join the values as strings with a hyphen
                        string result = string.Join("-", parameterPINs.Select(pin => pin.ToString()));

                        // Set the cell format to Text to prevent Excel from interpreting it as a date
                        ReportWS.Cells[ReportAppendRow, 5].NumberFormat = "@";  // Set to Text format

                        // Assign the result to the cell
                        ReportWS.Cells[ReportAppendRow, 5].Value = result;

                        // Set the font to bold
                        ReportWS.Cells[ReportAppendRow, 5].Font.Bold = true;

                        ReportAppendRow += 2;
                        parameterPINs.Clear();
                    }
                    ConnectorNameandPinTER = string.Empty;
                    currentConnectorName = string.Empty;
                }

                ReportWS.Columns[1].AutoFit();
                ReportWB.Save();
                return ReportWS.Name;
            }
            catch (Exception ex)
            {
                Logging.Error("Error while appending PowerOn data: " + ex.Message);
                return "null";
            }
        }*/

        // Commented on June 2nd before implementing JM multiple destinations
        /* public string AppendToExcelPowerOn(List<ElectreObject> elecCollection)
         {
             try
             {
                 // Step 1: Get only SCBs & TCBs as source components
                 var sourceComponents = GetSCB_TCB_Source(elecCollection);
                 if (sourceComponents.Count == 0)
                 {
                     MessageBox.Show($"Circuit breakers not found in the Panel", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                     Logging.Info("Circuit breakers not found in the Panel");
                 }
                 string ConnectorNameandPinTER = string.Empty;
                 string breakConnectorName = string.Empty;

                 List<string> parameterPINs = new List<string>();

                 // Step 2: Loop through each SCB/TCB as a starting point
                 foreach (var source in sourceComponents)
                 {
                     string currentWireNumber = source.WireNumber;
                     string currentConnectorName = source.ConnectorName;
                     string currentPinNumber = source.PinNumber;
                     // string tag = source.Core_Part_Number;
                     string subNet = source.SubNet;

                     string groundConnectorName = string.Empty;
                     string groundPinNumber = string.Empty;

                     // Find the most relevant GND pin
                     string negPin = PowerOnReport.FindGroundPin(elecCollection, source);

                     if (!string.IsNullOrEmpty(negPin))
                     {
                         if (negPin.Contains(","))
                         {
                             string[] negPinParts = negPin.Split(',');
                             groundConnectorName = negPinParts[0];
                             groundPinNumber = negPinParts[1];
                         }
                     }

                     //  List<string> visited = new List<string>();
                     List<string> visitedConnections = new List<string>();

                     // Step 3: Trace the continuity path
                     while (!string.IsNullOrEmpty(currentWireNumber))
                     {
                         // visited.Add(visitKey);
                         visitedConnections.Add($"{currentConnectorName},{currentPinNumber}");

                         // Find all objects that share the same wire number and subnet
                         var connectedObjects = modMain.ElecCollection_All
                                                       .Where(w => w.WireNumber + "," + w.SubNet == currentWireNumber + "," + subNet &&
                                                       !visitedConnections.Contains($"{w.ConnectorName},{w.PinNumber}") && !string.IsNullOrEmpty(w.ConnectorName))
                                                       .ToList();

                         if (connectedObjects.Count == 0)
                             break; // Stop if no further connections or loop detected

                         bool foundNext = false;

                         foreach (var connObj in connectedObjects)
                         {
                             if (connObj.ConnectorName == currentConnectorName && connObj.PinNumber == currentPinNumber)
                                 continue; // Skip self

                             // Update last known connector details
                             currentConnectorName = connObj.ConnectorName;
                             currentPinNumber = connObj.PinNumber;


                             visitedConnections.Add($"{currentConnectorName},{currentPinNumber}");

                             switch (connObj.ComponentType)
                             {
                                 case "EQU":
                                     currentConnectorName = $"{connObj.ConnectorName}";
                                     if (currentConnectorName.EndsWith("M", StringComparison.OrdinalIgnoreCase) || currentConnectorName.EndsWith("F", StringComparison.OrdinalIgnoreCase))
                                     {
                                         breakConnectorName = PowerOnReport.TracePinofBreakConnectorGND(connObj);
                                         if (string.IsNullOrEmpty(breakConnectorName))
                                         {
                                             breakConnectorName = "null";
                                             break;
                                         }
                                         else
                                         {
                                             currentConnectorName = breakConnectorName;
                                         }
                                     }
                                     break;
                                 case "DIS":
                                     breakConnectorName = PowerOnReport.TracePinofBreakConnectorGND(connObj);
                                     if (string.IsNullOrEmpty(breakConnectorName))
                                     {
                                         breakConnectorName = "null";
                                         break;
                                     }
                                     else
                                     {
                                         currentConnectorName = breakConnectorName;
                                     }
                                     break;
                                 case "TBK":
                                     var ConnectorNameandPinTBK = PowerOnReport.TracePinOfJMGND(connObj);
                                     if (!string.IsNullOrEmpty(ConnectorNameandPinTBK))
                                     {
                                         string[] parts = ConnectorNameandPinTBK.Split(',');
                                         currentConnectorName = parts[0];
                                         currentPinNumber = parts[1];
                                     }
                                     break;
                                 case "SPL":
                                     var ConnectorNameandPinSPL = PowerOnReport.TracePinOfSPLGND(connObj);
                                     if (!string.IsNullOrEmpty(ConnectorNameandPinSPL))
                                     {
                                         string[] parts = ConnectorNameandPinSPL.Split(',');
                                         currentConnectorName = parts[0];
                                         currentPinNumber = parts[1];
                                     }
                                     break;
                                 case "TER":

                                     ConnectorNameandPinTER = PowerOnReport.TracePinOfTERGND(connObj);
                                     if (!string.IsNullOrEmpty(ConnectorNameandPinTER))
                                     {
                                         string[] parts = ConnectorNameandPinTER.Split(',');
                                         currentConnectorName = parts[0];
                                         currentPinNumber = parts[1];
                                     }
                                     else if (string.IsNullOrEmpty(ConnectorNameandPinTER))
                                     {
                                         ConnectorNameandPinTER = "null";
                                     }
                                     break;
                                 case "REL":
                                 case "SWT":
                                 case "DD":

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
                                 default:
                                     currentConnectorName = connObj.ConnectorName;
                                     break;//pass the for the empty value
                             }
                             if (ConnectorNameandPinTER == "null")
                             {
                                 break;
                             }
                             if (breakConnectorName == "null")
                             {
                                 break;
                             }

                             // Step 4: Find the next wire connection based on the new connector
                             var nextWireConnection = modMain.ElecCollection_All
                                 .FirstOrDefault(w => string.Equals(w.ConnectorName, currentConnectorName, StringComparison.OrdinalIgnoreCase) && w.PinNumber == currentPinNumber);

                             if (nextWireConnection != null)
                             {
                                 // Move to the next wire
                                 currentWireNumber = nextWireConnection.WireNumber;
                                 currentConnectorName = nextWireConnection.ConnectorName;
                                 currentPinNumber = nextWireConnection.PinNumber;
                                 // tag = nextWireConnection.Core_Part_Number;
                                 subNet = nextWireConnection.SubNet;

                                 foundNext = true;
                                 break; // Move to next wire
                             }
                         }

                         if (!foundNext)
                             break; // Stop if no further valid connections
                     }
                     ReportWS.Cells[ReportAppendRow, 1].Value = source.Panel;
                     ReportWS.Cells[ReportAppendRow, 2].Value = source.ConnectorName; // SCB/TCB Source
                     ReportWS.Cells[ReportAppendRow, 3].Value = currentConnectorName; // Final Destination Unit                
                     ReportWS.Cells[ReportAppendRow, 4].Value = $"{currentPinNumber}(+)"; // positive destination pin

                     if (string.IsNullOrEmpty(negPin))
                     {
                         ReportWS.Cells[ReportAppendRow + 1, 2].Value = "";
                         ReportWS.Cells[ReportAppendRow + 1, 3].Value = "wrt STR";
                         ReportWS.Cells[ReportAppendRow + 1, 4].Value = "-";
                     }
                     else if (!string.IsNullOrEmpty(negPin))
                     {
                         ReportWS.Cells[ReportAppendRow + 1, 2].Value = "";
                         ReportWS.Cells[ReportAppendRow + 1, 3].Value = $"{groundConnectorName}"; // Final Destination Unit for ground
                         ReportWS.Cells[ReportAppendRow + 1, 4].Value = $"{groundPinNumber}(-)";  // negative destination pin
                     }

                     ReportWS.Cells[ReportAppendRow, 5].Value = source.Voltage; // Parameter Column

                     ReportWS.Range[ReportWS.Cells[ReportAppendRow, 5], ReportWS.Cells[ReportAppendRow + 1, 5]].Merge();  // Merge the rows in column 5


                     // Merge the rows in column 1
                     if (string.Equals(source.ComponentType, "SCB", StringComparison.OrdinalIgnoreCase))
                     {
                         ReportWS.Range[ReportWS.Cells[ReportAppendRow, 1], ReportWS.Cells[ReportAppendRow + 1, 1]].Merge();
                     }
                     else if (string.Equals(source.ComponentType, "TCB", StringComparison.OrdinalIgnoreCase))
                     {
                         ReportWS.Range[ReportWS.Cells[ReportAppendRow, 1], ReportWS.Cells[ReportAppendRow + 3, 1]].Merge();
                     }

                     // Merge the rows in column 2
                     ReportWS.Range[ReportWS.Cells[ReportAppendRow, 2], ReportWS.Cells[ReportAppendRow + 1, 2]].Merge();

                     // ReportWS.Rows.AutoFit();  // Auto-size column width
                     ReportWS.Rows[ReportAppendRow].AutoFit();
                     ReportWS.Rows[ReportAppendRow + 1].AutoFit();

                     // Auto-fit for 4th Column "Parameter"
                     // ReportWS.Columns[4].AutoFit();

                     ReportAppendRow += 3;

                     if (string.Equals(source.ComponentType, "TCB", StringComparison.OrdinalIgnoreCase))
                     {
                         // adds if component type is TCB only
                         parameterPINs.Add(currentPinNumber);
                     }

                     if (parameterPINs.Count == 3)
                     {
                         // Join the values as strings with a hyphen
                         string result = string.Join("-", parameterPINs.Select(pin => pin.ToString()));

                         // Set the cell format to Text to prevent Excel from interpreting it as a date
                         ReportWS.Cells[ReportAppendRow, 5].NumberFormat = "@";  // Set to Text format

                         // Assign the result to the cell
                         ReportWS.Cells[ReportAppendRow, 5].Value = result;

                         // Set the font to bold
                         ReportWS.Cells[ReportAppendRow, 5].Font.Bold = true;

                         ReportAppendRow += 2;
                         parameterPINs.Clear();
                     }
                     ConnectorNameandPinTER = string.Empty;
                     currentConnectorName = string.Empty;
                 }

                 ReportWS.Columns[1].AutoFit();
                 ReportWB.Save();
                 return ReportWS.Name;
             }
             catch (Exception ex)
             {
                 Logging.Error("Error while appending PowerOn data: " + ex.Message);
                 return "null";
             }
         }*/

        /*public void AppendToExcelPowerOnNew(List<ElectreObject> elecCollection)
        {
            // Step 1: Get only SCBs & TCBs as source components
            var sourceComponents = GetSCB_TCB_Source(elecCollection);
            if (sourceComponents.Count == 0)
            {
                MessageBox.Show($"Circuit breakers not found in the Panel", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Logging.Info("Circuit breakers not found in the Panel");
            }
            string ConnectorNameandPinTER = string.Empty;
            string breakConnectorName = string.Empty;

            List<string> parameterPINs = new List<string>();

            foreach (var source in sourceComponents)
            {
                string currentConnectorName = source.ConnectorName;
                string currentPinNumber = source.PinNumber;
                string currentWireNumber = source.WireNumber;
                string subNet = source.SubNet;

                string groundConnectorName = string.Empty;
                string groundPinNumber = string.Empty;

                // Find the most relevant GND pin
                string negPin = PowerOnReport.FindGroundPin(elecCollection, source);

                if (!string.IsNullOrEmpty(negPin))
                {
                    if (negPin.Contains(","))
                    {
                        string[] negPinParts = negPin.Split(',');
                        groundConnectorName = negPinParts[0];
                        groundPinNumber = negPinParts[1];
                    }
                }

                // Trace all destination TERs
                var destinations = TraceContinuityPath(currentConnectorName, currentPinNumber, currentWireNumber, subNet, new HashSet<string>());

                foreach (var (destConnector, destPin) in destinations)
                {
                    ReportWS.Cells[ReportAppendRow, 1].Value = source.Panel;
                    ReportWS.Cells[ReportAppendRow, 2].Value = source.ConnectorName;
                    ReportWS.Cells[ReportAppendRow, 3].Value = destConnector;
                    ReportWS.Cells[ReportAppendRow, 4].Value = $"{destPin}(+)";

                    if (string.IsNullOrEmpty(groundPinNumber))
                    {
                        ReportWS.Cells[ReportAppendRow + 1, 3].Value = "wrt STR";
                        ReportWS.Cells[ReportAppendRow + 1, 4].Value = "-";
                    }
                    else
                    {
                        ReportWS.Cells[ReportAppendRow + 1, 3].Value = groundConnectorName;
                        ReportWS.Cells[ReportAppendRow + 1, 4].Value = $"{groundPinNumber}(-)";
                    }

                    ReportWS.Cells[ReportAppendRow, 5].Value = source.Voltage;
                    ReportWS.Range[ReportWS.Cells[ReportAppendRow, 5], ReportWS.Cells[ReportAppendRow + 1, 5]].Merge();
                    ReportWS.Range[ReportWS.Cells[ReportAppendRow, 2], ReportWS.Cells[ReportAppendRow + 1, 2]].Merge();

                    if (string.Equals(source.ComponentType, "SCB", StringComparison.OrdinalIgnoreCase))
                    {
                        ReportWS.Range[ReportWS.Cells[ReportAppendRow, 1], ReportWS.Cells[ReportAppendRow + 1, 1]].Merge();
                    }
                    else if (string.Equals(source.ComponentType, "TCB", StringComparison.OrdinalIgnoreCase))
                    {
                        ReportWS.Range[ReportWS.Cells[ReportAppendRow, 1], ReportWS.Cells[ReportAppendRow + 3, 1]].Merge();
                    }

                    ReportAppendRow += 3;

                    if (string.Equals(source.ComponentType, "TCB", StringComparison.OrdinalIgnoreCase))
                    {
                        // adds if component type is TCB only
                        parameterPINs.Add(currentPinNumber);
                    }

                    if (parameterPINs.Count == 3)
                    {
                        // Join the values as strings with a hyphen
                        string result = string.Join("-", parameterPINs.Select(pin => pin.ToString()));

                        // Set the cell format to Text to prevent Excel from interpreting it as a date
                        ReportWS.Cells[ReportAppendRow, 5].NumberFormat = "@";  // Set to Text format

                        // Assign the result to the cell
                        ReportWS.Cells[ReportAppendRow, 5].Value = result;

                        // Set the font to bold
                        ReportWS.Cells[ReportAppendRow, 5].Font.Bold = true;

                        ReportAppendRow += 2;
                        parameterPINs.Clear();
                    }
                    ReportWB.Save();
                }
            }
        }*/

        /*  private List<(string Connector, string Pin)> TraceContinuityPath(string startConnector, string startPin, string wireNumber, string subNet, HashSet<string> visited)
          {
              List<(string Connector, string Pin)> endpoints = new();

              // Prevent infinite loop
              if (visited.Contains($"{startConnector},{startPin}"))
                  return endpoints;

              visited.Add($"{startConnector},{startPin}");

              var connectedObjects = modMain.ElecCollection_All
                  .Where(w => w.WireNumber == wireNumber && w.SubNet == subNet &&
                              !(w.ConnectorName == startConnector && w.PinNumber == startPin))
                  .ToList();

              foreach (var connObj in connectedObjects)
              {
                  string nextConnector = connObj.ConnectorName;
                  string nextPin = connObj.PinNumber;

                  switch (connObj.ComponentType)
                  {
                      case "TBK":
                          var tbkPins = PowerOnReport.TracePinOfJM(connObj);
                          foreach (var pinStr in tbkPins)
                          {
                              var parts = pinStr.Split(',');
                              var c = parts[0];
                              var p = parts[1];
                              endpoints.AddRange(TraceContinuityPath(c, p, connObj.WireNumber, connObj.SubNet, new HashSet<string>(visited)));
                          }
                          break;

                      case "SPL":
                          var splPins = PowerOnReport.TracePinOfSPL(connObj);
                          foreach (var pinStr in splPins)
                          {
                              var parts = pinStr.Split(',');
                              var c = parts[0];
                              var p = parts[1];
                              endpoints.AddRange(TraceContinuityPath(c, p, connObj.WireNumber, connObj.SubNet, new HashSet<string>(visited)));
                          }
                          break;

                      case "TER":
                          var terResult = PowerOnReport.TracePinOfTER(connObj);
                          if (!string.IsNullOrEmpty(terResult) && terResult.Contains(","))
                          {
                              var parts = terResult.Split(',');
                              endpoints.AddRange(TraceContinuityPath(parts[0], parts[1], connObj.WireNumber, connObj.SubNet, new HashSet<string>(visited)));
                          }
                          else
                          {
                              endpoints.Add((connObj.ConnectorName, connObj.PinNumber));
                          }
                          break;

                      case "DIS":
                      case "EQU":
                          string breakConnectorName = null;
                          if (connObj.ConnectorName.EndsWith("M", StringComparison.OrdinalIgnoreCase) ||
                              connObj.ConnectorName.EndsWith("F", StringComparison.OrdinalIgnoreCase))
                          {
                              breakConnectorName = PowerOnReport.TracePinofBreakConnector(connObj);
                              if (string.IsNullOrEmpty(breakConnectorName))
                              {
                                  continue;
                              }
                          }
                          else
                          {
                              breakConnectorName = connObj.ConnectorName;
                          }

                          var nextWireConnection = modMain.ElecCollection_All
                                  .FirstOrDefault(w => string.Equals(w.ConnectorName, breakConnectorName, StringComparison.OrdinalIgnoreCase) && w.PinNumber == connObj.PinNumber);

                          if (nextWireConnection!=null)
                          {
                              endpoints.AddRange(TraceContinuityPath(nextWireConnection.ConnectorName, nextWireConnection.PinNumber, nextWireConnection.WireNumber, nextWireConnection.SubNet, new HashSet<string>(visited)));
                          }
                          break;

                      default:
                          endpoints.AddRange(TraceContinuityPath(nextConnector, nextPin, connObj.WireNumber, connObj.SubNet, new HashSet<string>(visited)));
                          break;
                  }
              }

              return endpoints;
          }*/

        #endregion

        public bool CreatePowerOnReportHeader(string workbookPath)
        {
            ReportAppendRow = 1;

            try
            {
                SuspendExcelUpdates();
                if (!IsWorkbookOpen(workbookPath))
                    ReportWB = ExcelApp.Workbooks.Open(workbookPath);

                ReportWS = ReportWB.ActiveSheet as Worksheet;

                if (ReportWS == null)
                {
                    Logging.Error("Failed to access the active sheet.");
                    return false;
                }

                object[,] header = new object[1, 6];
                header[0, 0] = "Distribution Box";
                header[0, 1] = "CB to be pressed";
                header[0, 2] = "Unit";
                header[0, 3] = "Pins";
                header[0, 4] = "Parameter";
                header[0, 5] = "Remarks (MeasuredVoltage)";
                WriteBlock(ReportWS, ReportAppendRow, 1, header);

                Range headerRange = ReportWS.Range["A1", "F1"];
                headerRange.Font.Bold = true;
                headerRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;

                ((Range)ReportWS.Columns["A"]).ColumnWidth = 20;
                ((Range)ReportWS.Columns["B"]).ColumnWidth = 20;
                ((Range)ReportWS.Columns["C"]).ColumnWidth = 15;
                ((Range)ReportWS.Columns["D"]).ColumnWidth = 15;
                ((Range)ReportWS.Columns["E"]).ColumnWidth = 20;
                ((Range)ReportWS.Columns["F"]).ColumnWidth = 30;

                // Same centering as A2:F{Rows.Count}, without formatting every unused row.
                Range dataCols = ReportWS.Range["A:F"];
                dataCols.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                dataCols.VerticalAlignment = XlHAlign.xlHAlignCenter;
                ReleaseCom(dataCols);
                ReleaseCom(headerRange);

                ReportAppendRow++;

                Logging.Info($"wireList report Header created for {workbookPath}");

                /*  string filePath = Path.Combine(frmMain.lblTempFolder.Text, "PowerON.txt");
                  int ReportRows = ReadTextToCurrentExcel(filePath, ";", 2, 2,ReportWS);
                  MergeCBinReport(ReportRows, 2, 2);*/

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Excel ##06: " + ex.Message);
                Logging.Error("Failed to create WireLess report Header: " + ex.Message);
                return false;
            }
        }

        private bool IsWorkbookOpen(string workbookPath)
        {
            if (ReportWB == null || string.IsNullOrEmpty(workbookPath))
                return false;
            try
            {
                string openPath = Path.GetFullPath(ReportWB.FullName);
                string want = Path.GetFullPath(workbookPath);
                return string.Equals(openPath, want, StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return false;
            }
        }
       
        //this method converts data from 2d array to Dictionary for megger sheet3
        public  Dictionary<string, List<string>> Convert2DArrayToDictionary(object[,] array2D)
       {
            Dictionary<string, List<string>> dict = new Dictionary<string, List<string>>();

            for (int i = 0; i < array2D.GetLength(0); i++)
            {
                string key = array2D[i, 0]?.ToString(); // First column as key
                List<string> values = new List<string>();

                for (int j = 1; j < array2D.GetLength(1); j++) // Remaining columns as values
                {
                    if (array2D[i, j] != null) // Ignore null values
                    {
                        values.Add(array2D[i, j].ToString());
                    }
                }

                dict[key] = values;
            }
            return dict;
       }

        private void DeleteDefaultSheets(Excel.Workbook wb)
        {
            foreach (Excel.Worksheet sheet in wb.Sheets)
            {
                // Only delete if not your copied sheets (optional condition)
                if (!string.Equals(sheet.Name, "MS", StringComparison.OrdinalIgnoreCase)
                    && !sheet.Name.StartsWith("ML-", StringComparison.OrdinalIgnoreCase))
                {
                    sheet.Delete();
                }
            }
        }

        public bool ReadOOTBMaterialListXL()
        {
            var projectPath = Path.GetDirectoryName(GlobalVar.ReportFolderGlobal.TrimEnd('\\')) + "\\";
            string projectName = new DirectoryInfo(projectPath.TrimEnd('\\')).Name;
            Excel.Workbook ootbWBms = null;
            try
            {
                SuspendExcelUpdates();
                string ootbWBms_xls = projectPath + "result\\" + projectName + "_mat.xls";
                string ootbWBms_xlsx = projectPath + "result\\" + projectName + "_mat.xlsx";

                // Check for file existence and open the appropriate one
                if (!File.Exists(ootbWBms_xls))
                {
                    if (!File.Exists(ootbWBms_xlsx))
                    {
                        Console.WriteLine("The OOTB Material list file is missing");
                        Console.WriteLine("MS Report is not generated.");
                        return false;
                    }
                    else
                    {
                        ootbWBms = ExcelApp.Workbooks.Open(ootbWBms_xlsx);
                    }
                }
                else
                {
                    ootbWBms = ExcelApp.Workbooks.Open(ootbWBms_xls);
                }

                // Read panel sheet
                if (string.IsNullOrEmpty(modMain.ChkListPanelText))
                {
                    Console.WriteLine("Sheet name (modMain.ChkListPanelText) is not set.");
                    return false;
                }
                Excel.Worksheet OOTBmsSheet = ootbWBms.Sheets["F"] as Excel.Worksheet;

                if (OOTBmsSheet == null)
                {
                    Console.WriteLine("The Panel MS sheet is not in " + ootbWBms.Name);
                    Console.WriteLine("Panel MS is not generated.");
                    return false;
                }

                const int maxOotbRows = 3000;
                object[,] panelBlock = OOTBmsSheet.Range["A1:I" + maxOotbRows].Value2 as object[,];
                int i = 1;
                if (panelBlock != null)
                {
                    int blockRows = panelBlock.GetLength(0);
                    while (i <= blockRows && panelBlock[i, 1] != null)
                    {
                        for (int S = 1; S <= 9; S++)
                        {
                            var cellValue = panelBlock[i, S];
                            modMain.arrTableOfOOTBms_Panel[i, S] = cellValue != null ? cellValue.ToString() : string.Empty;
                        }
                        i++;
                    }
                }
                modMain.OOTBmsTotalrow_Selection = i;

                // Read weight from "WEIGHT RESULT" sheet
                OOTBmsSheet = ootbWBms.Sheets["WEIGHT RESULT"] as Excel.Worksheet;
                object[,] weightBlock = OOTBmsSheet.Range["A1:B200"].Value2 as object[,];
                if (weightBlock != null)
                {
                    int weightRows = weightBlock.GetLength(0);
                    for (int yy = 1; yy <= weightRows; yy++)
                    {
                        var panelName = weightBlock[yy, 1];
                        if (panelName != null && panelName.ToString() == modMain.ChkListPanelText)
                        {
                            var weightVal = weightBlock[yy, 2];
                            OOTBPaneltotalWeight = weightVal != null ? Convert.ToString(weightVal) : OOTBPaneltotalWeight;
                            break;
                        }
                    }
                }

                ootbWBms.Close(false);
                ReleaseCom(ootbWBms);
                ootbWBms = null;
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
                return false;
            }
            finally
            {
                if (ootbWBms != null)
                {
                    try { ootbWBms.Close(false); } catch { }
                    ReleaseCom(ootbWBms);
                }
            }
        }

        private void PrepareMaterialListWorkbookFromTemplate(int numberOfSheetsRequired, bool useXlsx, string reportFilePath)
        {
            ExcelApp.DisplayAlerts = false;
            Worksheet keep = FindSheetByName(ReportWB, "MS") ?? FindSheetByName(ReportWB, "ML-1");
            if (keep == null)
                throw new InvalidOperationException("Template sheet 'MS' not found.");

            string keepName = keep.Name;
            for (int i = ReportWB.Sheets.Count; i >= 1; i--)
            {
                Worksheet ws = (Worksheet)ReportWB.Sheets[i];
                if (!string.Equals(ws.Name, keepName, StringComparison.OrdinalIgnoreCase))
                    ws.Delete();
            }

            Worksheet first = (Worksheet)ReportWB.Sheets[1];
            if (first.Name != "ML-1")
                first.Name = "ML-1";

            first.Range["B8"].Value2 = PanelDrawingWindow.panelDigit;
            first.Range["H9"].Value2 = OOTBPaneltotalWeight;

            if (useXlsx)
                ReportWB.SaveAs(reportFilePath, XlFileFormat.xlOpenXMLWorkbook);

            DuplicateTemplateSheets(numberOfSheetsRequired, "ML-");
        }

        private void SaveMaterialListWorkbook(bool useXlsx, string reportFilePath)
        {
            if (useXlsx)
                ReportWB.Save();
            else
                ReportWB.SaveAs(reportFilePath, XlFileFormat.xlExcel8);
        }

        private void WriteMaterialListDataBlock(Worksheet ws, int startRow, int qStart, int rows)
        {
            object[,] colB = new object[rows, 1];
            object[,] colD = new object[rows, 1];
            object[,] colE = new object[rows, 1];
            object[,] colG = new object[rows, 1];
            object[,] colH = new object[rows, 1];
            object[,] colK = new object[rows, 1];

            for (int r = 0; r < rows; r++)
            {
                int q = qStart + r;
                colB[r, 0] = q;
                colD[r, 0] = modMain.arrTableOfOOTBms_Panel[q + 1, 4];
                colE[r, 0] = modMain.arrTableOfOOTBms_Panel[q + 1, 5];
                colG[r, 0] = modMain.arrTableOfOOTBms_Panel[q + 1, 7];
                colH[r, 0] = modMain.arrTableOfOOTBms_Panel[q + 1, 9];

                string ak1 = modMain.arrTableOfOOTBms_Panel[q + 1, 2];
                if (ak1.StartsWith("LOC-", StringComparison.OrdinalIgnoreCase))
                    colK[r, 0] = ak1.Substring(4);
                else
                    colK[r, 0] = ak1;
            }

            // Same cells as the original per-cell writes: B, D, E, G, H, K (StartCol=2).
            WriteBlock(ws, startRow, 2, colB);
            WriteBlock(ws, startRow, 4, colD);
            WriteBlock(ws, startRow, 5, colE);
            WriteBlock(ws, startRow, 7, colG);
            WriteBlock(ws, startRow, 8, colH);
            WriteBlock(ws, startRow, 11, colK);
        }

        public void GenerateHALReportFormat_MaterialList(string[] iarr, string iReportName)
        {
            string workingCopy = null;
            var sw = Stopwatch.StartNew();
            try
            {
                if (ExcelApp == null)
                    InitiateExcel();
                SuspendExcelUpdates();

                if (!ReadOOTBMaterialListXL()) return; // Check if the OOTB material list is created

                int NumRowInReportSheet = 16; // Number of rows per sheet in the report
                int StartRow = 13; // Start Row for data
                int NumberOfSheetsRequired = (int)Math.Floor((double)modMain.OOTBmsTotalrow_Selection / NumRowInReportSheet + 1); // Calculate the number of sheets needed

                string templatePath = Environment.GetEnvironmentVariable("ELECTRE_CUSTOMIZE") + "\\system\\MATERIALLIST_ReportFormat.xls";
                if (!File.Exists(templatePath))
                {
                    Console.WriteLine("The MaterialList Report Template is missing. So saving file into TEMPFILES folder without template");
                    return;
                }

                string panelDrawingPath = ConfigurationManager.AppSettings["PanelDrawingFolder"];
                string templFolder = Path.Combine(GlobalVar.ReportFolderGlobal, panelDrawingPath);

                if (NumberOfSheetsRequired < 1)
                    NumberOfSheetsRequired = 1;

                bool useXlsx = NumberOfSheetsRequired > 250;
                string reportFilePath = Path.Combine(templFolder, iReportName);
                if (useXlsx)
                {
                    reportFilePath = Path.Combine(templFolder, iReportName + ".xlsx");
                    Logging.Info($"Material list '{iReportName}' has {NumberOfSheetsRequired} sheets; saving as .xlsx because .xls is limited to 255 sheets.");
                }

                workingCopy = Path.Combine(templFolder, iReportName + "_ml.tmp.xls");
                if (File.Exists(workingCopy))
                    File.Delete(workingCopy);
                File.Copy(templatePath, workingCopy, true);

                ReportWB = ExcelApp.Workbooks.Open(workingCopy);
                PrepareMaterialListWorkbookFromTemplate(NumberOfSheetsRequired, useXlsx, reportFilePath);
                long copyMs = sw.ElapsedMilliseconds;
                Logging.Info($"Material list '{iReportName}': copied {NumberOfSheetsRequired} sheets in {copyMs} ms");

                int Q = 1;
                int lastQ = modMain.OOTBmsTotalrow_Selection - 2;
                if (lastQ < 1)
                    lastQ = 1;

                for (int S = 1; S <= NumberOfSheetsRequired; S++)
                {
                    Worksheet WS1 = (Worksheet)ReportWB.Sheets[S];
                    WS1.Range["B9"].Value2 = $"SHEET {S} OF {NumberOfSheetsRequired} SHEETS";

                    int rowsThisSheet = Math.Min(NumRowInReportSheet, lastQ - Q + 1);
                    if (rowsThisSheet > 0)
                    {
                        WriteMaterialListDataBlock(WS1, StartRow, Q, rowsThisSheet);
                        Q += rowsThisSheet;
                    }

                    if (Q > lastQ)
                    {
                        SaveMaterialListWorkbook(useXlsx, reportFilePath);
                        Logging.Info($"Material list '{iReportName}': copy={copyMs}ms fill={sw.ElapsedMilliseconds - copyMs}ms total={sw.ElapsedMilliseconds}ms sheets={NumberOfSheetsRequired}");
                        return;
                    }
                }

                SaveMaterialListWorkbook(useXlsx, reportFilePath);
                Logging.Info($"Material list '{iReportName}': copy={copyMs}ms fill={sw.ElapsedMilliseconds - copyMs}ms total={sw.ElapsedMilliseconds}ms sheets={NumberOfSheetsRequired}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
                Logging.Error("GenerateHALReportFormat_MaterialList: " + ex.Message);
            }
            finally
            {
                try
                {
                    if (TempWB != null)
                    {
                        try { TempWB.Close(false); } catch { }
                        Marshal.ReleaseComObject(TempWB);
                        TempWB = null;
                    }
                    if (ReportWB != null)
                    {
                        ReportWB.Close(true);
                        Marshal.ReleaseComObject(ReportWB);
                        ReportWB = null;
                    }
                    if (workingCopy != null && workingCopy.EndsWith("_ml.tmp.xls", StringComparison.OrdinalIgnoreCase)
                        && File.Exists(workingCopy))
                    {
                        try { File.Delete(workingCopy); } catch { }
                    }
                }
                catch (Exception cleanupEx)
                {
                    Console.WriteLine("Error during Excel cleanup: " + cleanupEx.Message);
                }
            }
        }

        public void CCasCSV(object[,] iarr, string iName)
        {
            string panelDrawingPath = ConfigurationManager.AppSettings["PanelDrawingFolder"];
            string filePath = Path.Combine("templ",panelDrawingPath, iName);
            if (!File.Exists(filePath)) 
            {
                using (File.Create(filePath)) { }
            }

            using (StreamWriter sw = new StreamWriter(filePath, append: true))
            {
                int rowCount = iarr.GetLength(0); // Number of rows
                int colCount = 5; // Fixed to 5 columns as in VB6

                for (int n = 0; n < rowCount; n++)
                {
                    string[] tempArr = new string[5];

                    for (int m = 0; m < colCount; m++)
                    {
                        object cellValue = iarr[n, m];
                        tempArr[m - 1] = cellValue?.ToString() ?? "";
                    }

                    if (!string.IsNullOrEmpty(tempArr[0]))
                    {
                        sw.WriteLine(string.Join(",", tempArr));
                    }
                    else
                    {
                        break;
                    }
                }
            }
        }

        // old code commented on may 8th 25
        /*  public void SortWirelistComponent(string sheetName, int rowStart, int rowEnd,modExcel modExcelInst)
          {
              try
              {
                  // Initialize the range and columns to sort
                  string sRange = "A" + rowStart + ":I" + rowEnd;
                  string colFilter1 = "A" + rowStart + ":A" + rowEnd;
                  string colFilter2 = "B" + rowStart + ":B" + rowEnd;

                  // Activate the worksheet
                  Excel.Worksheet sheet = (Excel.Worksheet)modExcelInst.ExcelApp.ActiveWorkbook.Worksheets[sheetName];
                  sheet.Activate();

                  // Clear any previous sorts
                  sheet.Sort.SortFields.Clear();

                  // Add the first sorting field (Column A)
                  sheet.Sort.SortFields.Add(Key: sheet.Range[colFilter1],
                      SortOn: Excel.XlSortOn.xlSortOnValues,
                      Order: Excel.XlSortOrder.xlAscending,
                      DataOption: Excel.XlSortDataOption.xlSortNormal);

                  // Add the second sorting field (Column B) with a custom sort order
                  sheet.Sort.SortFields.Add(Key: sheet.Range[colFilter2],
                      SortOn: Excel.XlSortOn.xlSortOnValues,
                      Order: Excel.XlSortOrder.xlAscending,
                      CustomOrder: "A,a,AA,aa,B,b,BB,bb",
                      DataOption: Excel.XlSortDataOption.xlSortNormal);

                  // Perform the sort operation
                  Excel.Sort sortObject = sheet.Sort;
                  sortObject.SetRange(sheet.Range[sRange]);
                  sortObject.Header = Excel.XlYesNoGuess.xlGuess;
                  sortObject.MatchCase = true;
                  sortObject.Orientation = Excel.XlSortOrientation.xlSortRows;
                  sortObject.SortMethod = Excel.XlSortMethod.xlPinYin;
                  sortObject.Apply();

                  // Clear any errors and reset selection
                  System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet); // Release the Excel sheet object

                  // Optional: Save the workbook if needed
                  modExcelInst.ExcelApp.ActiveWorkbook.Save();
              }
              catch (Exception ex)
              {
                  // Handle any errors
                  Console.WriteLine("An error occurred: " + ex.Message);
              }
          }*/

        public void SortWirelistComponent(string sheetName, int rowStart, int rowEnd, modExcel modExcelInst)
        {
            Excel.Worksheet sheet = null;
            Excel.SortFields sortFields = null;
            Excel.Sort sortObject = null;

            try
            {
                // Build range strings
                string sRange = $"A{rowStart}:I{rowEnd}";
                string colFilter1 = $"A{rowStart}:A{rowEnd}";
                string colFilter2 = $"B{rowStart}:B{rowEnd}";

                // Get worksheet reference
                sheet = (Excel.Worksheet)modExcelInst.ExcelApp.ActiveWorkbook.Worksheets[sheetName];

                // Get Sort object
                sortFields = sheet.Sort.SortFields;
                sortFields.Clear();

                // Add sort fields
                sortFields.Add(
                    Key: sheet.Range[colFilter1],
                    SortOn: Excel.XlSortOn.xlSortOnValues,
                    Order: Excel.XlSortOrder.xlAscending,
                    DataOption: Excel.XlSortDataOption.xlSortNormal);

                sortFields.Add(
                    Key: sheet.Range[colFilter2],
                    SortOn: Excel.XlSortOn.xlSortOnValues,
                    Order: Excel.XlSortOrder.xlAscending,
                    CustomOrder: "A,a,AA,aa,B,b,BB,bb",
                    DataOption: Excel.XlSortDataOption.xlSortNormal);

                // Configure sort
                sortObject = sheet.Sort;
                sortObject.SetRange(sheet.Range[sRange]);
                sortObject.Header = Excel.XlYesNoGuess.xlGuess;
                sortObject.MatchCase = true;
                sortObject.Orientation = Excel.XlSortOrientation.xlSortRows;
                sortObject.SortMethod = Excel.XlSortMethod.xlPinYin;
                sortObject.Apply();

                // Save workbook
                modExcelInst.ExcelApp.ActiveWorkbook.Save();
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during sorting: " + ex.Message);
            }
            finally
            {
                // Proper COM cleanup
                if (sortFields != null) Marshal.ReleaseComObject(sortFields);
                if (sortObject != null) Marshal.ReleaseComObject(sortObject);
                if (sheet != null) Marshal.ReleaseComObject(sheet);

                sortFields = null;
                sortObject = null;
                sheet = null;

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        public string[] ReadTextTo1DArray(string fileFullPath)
        {
            List<string> lines = new List<string>();

            try
            {
                using (StreamReader sr = new StreamReader(fileFullPath))
                {
                    while (!sr.EndOfStream)
                    {
                        string line = sr.ReadLine();
                        lines.Add(line);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error reading file: " + ex.Message);
            }

            return lines.ToArray();
        }

        public void AppendMaterialListTxt(string templFolder,string ChkListPanelText, string filePath, List<string> materialsList)
        {
            using (StreamWriter writer = new StreamWriter(filePath, append: true))
            {
                foreach (string material in materialsList)
                {
                    writer.WriteLine(material);
                }
            }
        }

        /* public void AppendPanelDetailsTxt(string templFolder, string ChkListPanelText, string filePath)
         {
             var writtenPairs = new HashSet<string>();

             using (StreamWriter writer = new StreamWriter(filePath, append: true))
             {
                 List<ElectreObject> panelDetails = modMain.ElecCollection_All
                     .Where(e => e.Panel == ChkListPanelText)
                     .ToList();

                 foreach (var obj in panelDetails)
                 {
                     var sourceCon = obj.ConnectorName;
                     var sourcePin = obj.PinNumber;

                     var destinationObj = modMain.ElecCollection_All.Where(e => e.WireNumber == obj.WireNumber
                                                                 && e.Core_Part_Number == obj.Core_Part_Number
                                                                 && e.SubNet == obj.SubNet
                                                                 && e.ConnectorName != obj.ConnectorName)
                                                                 .FirstOrDefault();

                     if (destinationObj == null)
                         continue;

                     var destCon = destinationObj.ConnectorName;
                     var destPin = destinationObj.PinNumber;

                     // Create a normalized key for the pair
                     var pair1 = sourceCon + sourcePin;
                     var pair2 = destCon + destPin;

                     // Sort the pair alphabetically to treat reversed pairs the same
                     var key = string.Compare(pair1, pair2) < 0
                         ? pair1 + "|" + pair2
                         : pair2 + "|" + pair1;

                     // If this pair was already written, skip
                     if (writtenPairs.Contains(key))
                         continue;

                     // Mark the pair as written
                     writtenPairs.Add(key);

                     // Write the line
                     writer.WriteLine($"{sourceCon},{sourcePin},{destCon},{destPin},{obj.WireNumber}/{obj.Gauge.Replace("#", "")}/{obj.Core_Part_Number}");
                 }
             }
         }*/

        public void AppendPanelDetailsTxt(string templFolder, string ChkListPanelText, string filePath)
        {
            using (StreamWriter writer = new StreamWriter(filePath, append: true))
            {
                var allIndex = ElectreTraceIndex.Build(modMain.ElecCollection_All);
                var all = modMain.ElecCollection_All;
                var panelDetails = new List<ElectreObject>();
                for (int i = 0; i < all.Count; i++)
                {
                    var e = all[i];
                    if (string.Equals(e.Panel, ChkListPanelText, StringComparison.OrdinalIgnoreCase))
                        panelDetails.Add(e);
                }
                int missingDestCount = 0;
                foreach (var obj in panelDetails)
                {   
                    var sourceCon = obj.ConnectorName;
                    var sourcePin = obj.PinNumber;
                    ElectreObject destinationObj = null;
                    var mates = allIndex.ConnectedOnWire(obj.WireNumber, obj.SubNet);
                    for (int m = 0; m < mates.Count; m++)
                    {
                        var e = mates[m];
                        if (e.ConnectorName != obj.ConnectorName)
                        {
                            destinationObj = e;
                            break;
                        }
                    }
                    var destCon = string.Empty;
                    var destPin = string.Empty;

                    if (destinationObj != null)
                    {
                        destCon = destinationObj.ConnectorName;
                        destPin = destinationObj.PinNumber;                        
                    }
                    else
                    {
                        missingDestCount++;
                        Logging.Warning_PD($"No destination found for wire '{obj.WireNumber}' from '{sourceCon}' pin '{sourcePin}'");
                    }

                    string core = !string.IsNullOrEmpty(obj.Core_Part_Number) && obj.Core_Part_Number.Length <= 2
                            ? "/" + obj.Core_Part_Number
                            : string.Empty;


                    writer.WriteLine($"{sourceCon},{sourcePin},{destCon},{destPin},{obj.WireNumber}/{obj.Gauge.Replace("#", "")}{core}");

                    // writer.WriteLine($"{sourceCon},{sourcePin},{destCon},{destPin},{obj.WireNumber}/{obj.Gauge.Replace("#", "")}/{obj.Core_Part_Number}");

                }

                if(missingDestCount > 0)
                {
                    MessageBox.Show(
                        $"Missing destinations detected for panel '{ChkListPanelText}'.\n\n" +
                        $"Missing count: {missingDestCount}\n\n" +
                        $"Please check the log file for detailed information.",
                        "Missing Destinations",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning
                    );

                }
            }
        }

        public void AppendPanelInfoTxt(string templFolder, string ChkListPanelText)
        {
            using (StreamWriter writer = new StreamWriter(Path.Combine(templFolder, "PanelInfo.txt")))
            {
                writer.WriteLine($"{PanelDrawingWindow.panelDigit},{ChkListPanelText}");
            }
        }

        public void AppendAdminTxt(string templFolder, string ChkListPanelText, string filePath, List<string> materialsList)
        {
            using (StreamWriter writer = new StreamWriter(filePath, append: true))
            {
                foreach (string material in materialsList)
                {
                    writer.WriteLine(material);
                }
            }
        }

        public void CreateComponentContinuityReportHeader()
        {
            try
            {
                Excel.Worksheet ReportWS = ReportWB.ActiveSheet;
                int ReportAppendRow = 1;

                // Setting header values
                ReportWS.Cells[ReportAppendRow, 1] = "FROM CONN";
                ReportWS.Cells[ReportAppendRow, 2] = "FROM PIN";
                ReportWS.Cells[ReportAppendRow, 3] = "TO CONN";
                ReportWS.Cells[ReportAppendRow, 4] = "TO PIN";
                ReportWS.Cells[ReportAppendRow, 5] = "WIRE CODE";
                ReportWS.Cells[ReportAppendRow, 6] = "WIRE TYPE";

                // Adjust column widths
                ReportWS.Columns["A:A"].ColumnWidth = 14.78;
                ReportWS.Columns["B:B"].ColumnWidth = 14.78;
                ReportWS.Columns["C:C"].ColumnWidth = 14;
                ReportWS.Columns["D:D"].ColumnWidth = 16.67;
                ReportWS.Columns["E:E"].ColumnWidth = 21.33;
                ReportWS.Columns["F:F"].ColumnWidth = 14.11;

                // Apply bold font to header row
                Excel.Range headerRange = ReportWS.get_Range("A1", "F1");
                headerRange.Font.Bold = true;

                // Increment the row counter for the next data entry
                ReportAppendRow++;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error in CreateComponentContinuityReportHeader: " + ex.Message);
            }
        }


    }

}

