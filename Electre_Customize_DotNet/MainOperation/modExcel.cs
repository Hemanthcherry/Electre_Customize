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
using Electre_Customize_DotNet.Objects;
using Electre_Customize_DotNet.Logs;
using Electre_Customize_DotNet.Reports;
using static System.Runtime.InteropServices.JavaScript.JSType;
using Electre_Customize_DotNet.ReportUI;
using System.Configuration;
using System.IO;
using Electre_Customize_DotNet.Helpers.ExcelHelpers;
using Electre_Customize_DotNet.Helpers.Global;
using Electre_Customize_DotNet.Helpers.Megger;
using Electre_Customize_DotNet.Helpers.PowerOn;

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
        /// <summary>Set on the parallel Excel workers: several Excel processes open the same template at once, so it is opened read-only.</summary>
        public bool TemplateReadOnly;
        public static Dictionary<string, List<string>> matchedComponents = new Dictionary<string, List<string>>();
        // This method checks if the Excel instance is created or not. If the Excel instance is null, it tries to create a new Excel instance

        public void InitiateExcel()
        {
            try
            {
                // Try to get a running instance of Excel
                ExcelApp = GetRunningExcelApplication();
                if (ExcelApp == null)
                {
                    ExcelApp = new Application();
                }
                ExcelApp.Visible = false;
                //ExcelApp.Visible = true;

                // Perf: this ExcelApp is a dedicated headless instance for report generation
                // (created above via CoCreateInstance / new Application(), never the user's own
                // interactive Excel session, and always Quit() in releaseExcel()), so suppressing
                // screen repaint/alerts/events here cannot affect anything the user sees elsewhere.
                // Calculation mode is intentionally left Automatic: report templates use live
                // cross-sheet Formula/FormulaR1C1 references (e.g. page headers/footers), so
                // switching to manual recalculation requires a guaranteed full recalc before every
                // save to avoid stale values - handled together with the Save-call-site work later.
                ExcelApp.ScreenUpdating = false;
                ExcelApp.EnableEvents = false;
                ExcelApp.DisplayAlerts = false;

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error initiating Excel:" + ex.Message);
                Logging.Error("Error initiating Excel: " + ex.Message);

            }
        }
        // this method create and return Excel Instance
        private Application GetRunningExcelApplication()
        {
            try
            {
                // Try to get a running instance of Excel via ROT (Running Object Table)
                const string progId = "Excel.Application";
                Type excelType = Type.GetTypeFromProgID(progId);
                dynamic excelInstance = Activator.CreateInstance(excelType);
                return excelInstance as Application;
            }
            catch (COMException ex)
            {
                // No running instance found
                MessageBox.Show("Exce ##01: "+ ex.Message);
                return null;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error getting running Excel instance: " + ex.Message);
                return null;
            }
        }

        // this method creates Excel Workbook for megger
        public string CreateNewWorkbookforMeggerScheduler(string iWorkbookName, string iSheetName1, string iSheetName2, string iSheetName3, string iSheetName4, string iSheetName5, string iSheetName6, string iSheetName7, string reportType)
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

                ReportWB.Sheets.Add(After: ReportWB.Sheets[ReportWB.Sheets.Count]);
                Worksheet newSheet5 = (Worksheet)ReportWB.Sheets[ReportWB.Sheets.Count];
                newSheet5.Name = iSheetName5;

                ReportWB.Sheets.Add(After: ReportWB.Sheets[ReportWB.Sheets.Count]);
                Worksheet newSheet6 = (Worksheet)ReportWB.Sheets[ReportWB.Sheets.Count];
                newSheet6.Name = iSheetName6;

                ReportWB.Sheets.Add(After: ReportWB.Sheets[ReportWB.Sheets.Count]);
                Worksheet newSheet7 = (Worksheet)ReportWB.Sheets[ReportWB.Sheets.Count];
                newSheet7.Name = iSheetName7;

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

                        // Set header cells
                        ReportWS.Cells[ReportAppendRow, 1].Value = "Component Name";
                        ReportWS.Cells[ReportAppendRow, 2].Value = "Part Numbers";
                        // Format the header row
                        Range headerRange = ReportWS.Range["A1", "B1"];
                        headerRange.Font.Bold = true;

                        // Set column widths
                        ReportWS.Columns["A"].ColumnWidth = 20;
                        ReportWS.Columns["B"].ColumnWidth = 20;

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
                        // Set header cells
                        ReportWS1.Cells[ReportAppendRow1, 1].Value = "Connectors";
                        ReportWS1.Cells[ReportAppendRow1, 2].Value = "Pins";


                        // Format the header row
                        Range headerRange = ReportWS1.Range["A1", "B1"];
                        headerRange.Font.Bold = true;

                        // Set column widths
                        ReportWS1.Columns["A"].ColumnWidth = 20;
                        ReportWS1.Columns["B"].ColumnWidth = 15;
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

                        // Write text in the first merged cells
                        ReportWS2.Cells[ReportAppendRow2, 1].Value = "From"; // Column A
                        ReportWS2.Cells[ReportAppendRow2, 3].Value = "To";   // Column C

                        // Merge Column A and B for "From"
                        Range fromRange = ReportWS2.Range["A" + ReportAppendRow2, "B" + ReportAppendRow2];

                        // Merge Column C and D for "To"
                        Range toRange = ReportWS2.Range["C" + ReportAppendRow2, "D" + ReportAppendRow2];
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
                        ReportAppendRow2++;

                        ReportWS2.Cells[ReportAppendRow2, 1].Value = "Connector No";
                        ReportWS2.Cells[ReportAppendRow2, 2].Value = "Pin No";
                        ReportWS2.Cells[ReportAppendRow2, 3].Value = "Connector No";
                        ReportWS2.Cells[ReportAppendRow2, 4].Value = "Pin No";


                        // Format the header row
                        Range headerRange = ReportWS2.Range["A2", "D2"];
                        headerRange.Font.Bold = true;

                        // Set column widths for readability

                        ReportWS2.Columns["A"].ColumnWidth = 15; // Connector No (From)
                        ReportWS2.Columns["B"].ColumnWidth = 10; // Pin No (From)
                        ReportWS2.Columns["C"].ColumnWidth = 15; // Connector No (To)
                        ReportWS2.Columns["D"].ColumnWidth = 10; // Pin No (To)


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
                ReportWB = ExcelApp.Workbooks.Add();
                ExcelApp.DisplayAlerts = false;

                // Add a new sheet and name it
                ReportWB.Sheets.Add(After: ReportWB.Sheets[ReportWB.Sheets.Count]);
                Worksheet newSheet = (Worksheet)ReportWB.Sheets[ReportWB.Sheets.Count];
                newSheet.Name = iSheetName;

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
                ReportWB = ExcelApp.Workbooks.Open(workbookPath);

                ReportWS = ReportWB.ActiveSheet as Worksheet;

                if (ReportWS == null)
                {
                    Logging.Error("Failed to access the active sheet.");
                    return false;
                }

                // Set header cells
                ReportWS.Cells[ReportAppendRow, 1].Value = "FROM CONN";
                ReportWS.Cells[ReportAppendRow, 2].Value = "FROM PIN";
                ReportWS.Cells[ReportAppendRow, 3].Value = "TO CONN";
                ReportWS.Cells[ReportAppendRow, 4].Value = "TO PIN";
                ReportWS.Cells[ReportAppendRow, 5].Value = "WIRE CODE";
                ReportWS.Cells[ReportAppendRow, 6].Value = "WIRE TYPE";
                ReportWS.Cells[ReportAppendRow, 7].Value = "LENGTH";
                ReportWS.Cells[ReportAppendRow, 8].Value = "RD";
                ReportWS.Cells[ReportAppendRow, 9].Value = "LD";
                ReportWS.Cells[ReportAppendRow, 10].Value = "NERD";
                ReportWS.Cells[ReportAppendRow, 11].Value = "Group Number";
                //  ReportWS.Cells[ReportAppendRow, "Shunt"].Value = "Shunt"; 
                // Uncomment if Shunt is required

                // Format the header row
                Range headerRange = ReportWS.Range["A1", "K1"];
                headerRange.Font.Bold = true;

                // Set column widths
                ReportWS.Columns["A"].ColumnWidth = 15;
                ReportWS.Columns["B"].ColumnWidth = 15;
                ReportWS.Columns["C"].ColumnWidth = 15;
                ReportWS.Columns["D"].ColumnWidth = 15;
                ReportWS.Columns["E"].ColumnWidth = 20;
                ReportWS.Columns["F"].ColumnWidth = 10;
                ReportWS.Columns["G"].ColumnWidth = 10;
                ReportWS.Columns["H"].ColumnWidth = 20;
                ReportWS.Columns["I"].ColumnWidth = 15;
                ReportWS.Columns["J"].ColumnWidth = 10;
                ReportWS.Columns["K"].ColumnWidth = 15;
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
                ReportWB = ExcelApp.Workbooks.Open(workbookPath);

                ReportWS = ReportWB.ActiveSheet;

                if (ReportWS == null)
                {
                    MessageBox.Show("reportWS  null");
                    Logging.Error("Failed to access the active sheet.");
                    return false;
                }

                // Set header cells
                ReportWS.Cells[ReportAppendRow, 1].Value = "Component Name";
                ReportWS.Cells[ReportAppendRow, 2].Value = "Part Number";


                Range headerRange = ReportWS.Range["A1","B1"];
                headerRange.Font.Bold = true;

                // Set column widths
                ReportWS.Columns["A"].ColumnWidth = 20;
                ReportWS.Columns["B"].ColumnWidth = 20;

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

        #region adding data to excel

        // this method Append data to EXcel sheets for megger
        public string AppendToExcelMegger(Worksheet Reportws, object[,] iarr, int iTillColumn = 0)
        {
            try
            {
                Logging.Info($"Started writing data to sheet: {Reportws.Name}");

                if (iarr == null || iarr.Length == 0 || iarr.GetLength(0) == 0 || iarr.GetLength(1) == 0)
                {
                    Logging.Error("Data array is null or empty.");
                    return "null";
                }

                int row = 2; // Start after headers

                if (Reportws.Name == "Continuity Components")
                {
                    var matchedComponents = MeggerListConverters.ToDictionary(iarr);

                    // Every component's rows are contiguous (column 1 = component name on its first row, column 2 = one
                    // part number per row), and the merge ranges only cover a component's OWN column-1 rows, whose
                    // non-first cells are never written. So all the values can go to the sheet in one bulk write first,
                    // and the merges / alignment are applied afterwards - same final sheet, far fewer COM calls
                    // (it used to be ~3 calls per value plus ~8 per component).
                    int totalPartRows = 0;
                    foreach (var component in matchedComponents)
                    {
                        totalPartRows += component.Value.Count;
                    }

                    if (totalPartRows > 0)
                    {
                        object[,] block = new object[totalPartRows, 2];
                        var mergeGroups = new List<(int FirstRow, int PartCount)>();
                        int offset = 0;

                        foreach (var component in matchedComponents)
                        {
                            int partCount = component.Value.Count;

                            if (partCount > 0)
                            {
                                block[offset, 0] = component.Key;
                                for (int i = 0; i < partCount; i++)
                                {
                                    block[offset + i, 1] = component.Value[i];
                                }

                                mergeGroups.Add((row + offset, partCount));
                                offset += partCount;
                            }
                        }

                        ExcelBulk.WriteValues(Reportws, row, 1, block);

                        foreach (var group in mergeGroups)
                        {
                            var mergeRange = ExcelBulk.RangeAt(Reportws, group.FirstRow, 1, group.FirstRow + group.PartCount - 1, 1);
                            mergeRange.Merge();
                            mergeRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                            mergeRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                        }
                    }

                    ReportWB.Save();
                    Logging.Info("Data written to Continuity Components sheet.");
                    return Reportws.Name;
                }
                else if (Reportws.Name == "Pin List")
                {
                    int tillCol = iTillColumn != 0 ? iTillColumn : iarr.GetLength(1);
                    int totalRowsPL = iarr.GetLength(0);
                    int firstRowPL = row;

                    // First pass: replicate the original loop's break-before-increment bookkeeping
                    // exactly (including its existing off-by-one on the format range below when a
                    // blank-row break occurs) without touching Excel yet.
                    int rowsWrittenPL = 0;
                    for (int i = 0; i < totalRowsPL; i++)
                    {
                        rowsWrittenPL++;
                        if (totalRowsPL > 3 && ExcelRowRules.IsBlockBlank(iarr, i))
                        {
                            break;
                        }
                        row++;
                    }

                    if (rowsWrittenPL > 0 && tillCol > 0)
                    {
                        object[,] buffer = new object[rowsWrittenPL, tillCol];
                        for (int i = 0; i < rowsWrittenPL; i++)
                        {
                            for (int j = 0; j < tillCol; j++)
                            {
                                buffer[i, j] = iarr[i, j];
                            }
                        }

                        ExcelBulk.WriteValues(Reportws, firstRowPL, 1, buffer, centered: true);
                    }

                    Range range = Reportws.Range["A1", $"B{row - 1}"];
                    ExcelFormatting.FormatRange(range);
                    ReportWB.Save();

                    Logging.Info("Data written to Pin List sheet.");
                    return Reportws.Name;
                }

                else if (Reportws.Name == "Connection List")
                {
                    int tillCol = iTillColumn != 0 ? iTillColumn : iarr.GetLength(1);
                    int row1 = 3; // START from Row 3 (row 1 = merged header, row 2 = subheaders)
                    int totalRowsCL = iarr.GetLength(0);

                    // Only non-blank source rows get written (blank ones are skipped via `continue`
                    // in the original, so the destination rows are consecutive even though the
                    // source indices are not). The post-write break check in the original is
                    // unreachable here (same condition already causes `continue` above), so it is
                    // safely omitted rather than reproduced.
                    var rowIndexesCL = new List<int>();
                    for (int i = 0; i < totalRowsCL; i++)
                    {
                        if (ExcelRowRules.IsBlockBlank(iarr, i)) continue;
                        if (ExcelRowRules.IsRowBlank(iarr, i)) continue;
                        rowIndexesCL.Add(i);
                    }

                    if (rowIndexesCL.Count > 0 && tillCol > 0)
                    {
                        object[,] buffer = new object[rowIndexesCL.Count, tillCol];
                        for (int r = 0; r < rowIndexesCL.Count; r++)
                        {
                            for (int j = 0; j < tillCol; j++)
                            {
                                buffer[r, j] = iarr[rowIndexesCL[r], j];
                            }
                        }

                        ExcelBulk.WriteValues(Reportws, row1, 1, buffer, centered: true);
                        row1 += rowIndexesCL.Count;
                    }

                    // NOTE: uses the outer `row` (still its initial value here), exactly as the
                    // original code did - not row1. Preserved as-is, not "fixed".
                    Range range = Reportws.Range["A1", $"D{row - 1}"];
                    ExcelFormatting.FormatRange(range);
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
                    int totalRowsEx = iarr.GetLength(0);

                    // Break is checked before any write for the row, so the blank-triggering row
                    // itself is never written (unlike the Pin List branch above).
                    int rowsWrittenEx = totalRowsEx;
                    if (totalRowsEx > 3)
                    {
                        for (int i = 0; i < totalRowsEx; i++)
                        {
                            if (ExcelRowRules.IsBlockBlank(iarr, i))
                            {
                                rowsWrittenEx = i;
                                break;
                            }
                        }
                    }

                    if (rowsWrittenEx > 0)
                    {
                        object[,] buffer = new object[rowsWrittenEx, 3];
                        for (int i = 0; i < rowsWrittenEx; i++)
                        {
                            buffer[i, 0] = iarr[i, 0]; // Column A - Serial Number
                            buffer[i, 1] = iarr[i, 1]; // Column B - ConnectorName
                            buffer[i, 2] = iarr[i, 2]; // Column C - PinNumber
                        }

                        ExcelBulk.WriteValues(Reportws, row, 1, buffer, centered: true);
                        row += rowsWrittenEx;
                    }

                    Range range = Reportws.Range["A1", $"C{row - 1}"];
                    ExcelFormatting.FormatRange(range);

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
                            bool isBlankRow = ExcelRowRules.IsBlockBlank(iarr, k);
                            if (isBlankRow)
                            {
                                break;
                            }
                        }
                    }

                    // Format the range in Excel
                    Range range = ReportWS.Range["A1", $"B{ReportAppendRow - 1}"];
                    ExcelFormatting.FormatRange(range);


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
                            bool isBlankRow = ExcelRowRules.IsBlockBlank(iarr, k);
                            if (isBlankRow)
                            {
                                break;
                            }
                        }
                    }

                    // Format the range in Excel
                    Range range = ReportWS1.Range["A1", $"D{ReportAppendRow - 1}"];
                    ExcelFormatting.FormatRange(range);


                    return ReportWS1.Name;


                }
                else if(Reportws.Name == "Continuity Components")
                {
                   matchedComponents = MeggerListConverters.ToDictionary(iarr);

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
        public string AppendToExcel(object[,] iarr, int iTillColumn = 0, bool saveWorkbook = true)
        {
            try
            {
                int tillCol = iTillColumn != 0 ? iTillColumn : iarr.GetLength(1); // Determine the number of columns
                int totalRows = iarr.GetLength(0);

                // Determine how many rows are actually written, preserving the exact original
                // break-on-blank-row rule (only checked when the array has more than 3 rows).
                int rowsToWrite = totalRows;
                if (totalRows > 3)
                {
                    for (int k = 0; k < totalRows; k++)
                    {
                        if (ExcelRowRules.IsBlockBlank(iarr, k))
                        {
                            rowsToWrite = k + 1; // the blank-triggering row is still written, same as before
                            break;
                        }
                    }
                }

                if (rowsToWrite > 0 && tillCol > 0)
                {
                    object[,] buffer = new object[rowsToWrite, tillCol];
                    for (int k = 0; k < rowsToWrite; k++)
                    {
                        for (int j = 0; j < tillCol; j++)
                        {
                            buffer[k, j] = iarr[k, j];
                        }
                    }

                    ExcelBulk.WriteValues(ReportWS, ReportAppendRow, 1, buffer); // single bulk COM write instead of one call per cell
                }

                ReportAppendRow += rowsToWrite; // Move past the rows written, exactly as the original loop did

                // Format the range in Excel
                Range range = ReportWS.Range["A1", $"J{ReportAppendRow - 1}"];
                ExcelFormatting.FormatRange(range);

                // Save the workbook (callers that close the workbook with Close(true) right afterwards - the
                // sheet / component / project wire-list reports - pass false: the close saves the same content,
                // so this save was a redundant second disk write of the .xls)
                if (saveWorkbook)
                {
                    ReportWB.Save();
                }

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
                int row = 2; // Starting from row 2, as row 1 contains headers ("Component Name" and "Part Number")

                // Loop through the dictionary and add component name and part number
                foreach (var component in matchedComponents)
                {
                    int partCount = component.Value.Count;

                    // If there are part numbers, populate the first row with the component name and merge cells
                    if (partCount > 0)
                    {
                        // Populate "Component Name" in column 1 for the first part number
                        // (ExcelBulk.RangeAt = one COM call per cell/range; Cells[r, c] and Range[Cells, Cells] cost two and five)
                        ExcelBulk.RangeAt(ReportWS, row, 1, row, 1).Value = component.Key;

                        // Center align the text both horizontally and vertically for the merged cells in column 1
                        var componentCell = ExcelBulk.RangeAt(ReportWS, row, 1, row, 1);
                        componentCell.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        componentCell.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;

                        // Populate "Part Number" in column 2 for the first part number
                        ExcelBulk.RangeAt(ReportWS, row, 2, row, 2).Value = component.Value[0];

                        // Merge the first column for the number of rows corresponding to part numbers
                        ExcelBulk.RangeAt(ReportWS, row, 1, row + partCount - 1, 1).Merge();

                        // Loop through the remaining part numbers and populate column 2 (Part Numbers)
                        for (int i = 1; i < partCount; i++)
                        {
                            row++; // Move to the next row for the next part number

                            // Leave the first column empty (it is merged, and should remain the same for this component)
                            ExcelBulk.RangeAt(ReportWS, row, 1, row, 1).Value = "";

                            // Populate "Part Number" in column 2
                            ExcelBulk.RangeAt(ReportWS, row, 2, row, 2).Value = component.Value[i];
                        }

                        // Move to the next row after all part numbers for the current component are processed
                        row++;
                    }
                }
                // Save the workbook
                ReportWB.Save();

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
                // Close and release the workbook object
                if (ReportWB != null)
                {
                    ReportWB.Close(true);
                    Marshal.ReleaseComObject(ReportWB);
                }
            }
        }

        #endregion
        // header cells, formatting and column widths of the wire-list workbooks; leaves ReportAppendRow on the first data row
        private void WriteWirelistHeaderRows()
        {
            // Set header cells (one row write instead of ten single-cell writes)
            ExcelBulk.WriteValues(ReportWS, ReportAppendRow, 1, new object[,]
            {
                { "FROM CONN", "FROM PIN", "TO CONN", "TO PIN", "WIRE CODE", "WIRE TYPE", "LENGTH", "RD", "LD", " NERD" }
            });
           
            //  ReportWS.Cells[ReportAppendRow, "Shunt"].Value = "Shunt"; 
            // Uncomment if Shunt is required

            // Format the header row
            Range headerRange = ReportWS.Range["A1", "J1"];
            headerRange.Font.Bold = true;


            // Set column widths
            Range columns = ReportWS.Columns;
            columns["A"].ColumnWidth = 14.78;
            columns["B"].ColumnWidth = 14.78;
            columns["C"].ColumnWidth = 14;
            columns["D"].ColumnWidth = 16.67;
            columns["E"].ColumnWidth = 21.33;
            columns["F"].ColumnWidth = 14.11;
            columns["G"].ColumnWidth = 10;
            columns["H"].ColumnWidth = 20;
            columns["I"].ColumnWidth = 15;
            columns["J"].ColumnWidth = 10;
            
            ReportAppendRow++;
        }

        /// <summary>
        /// Opt-in (App.config OnePassWirelistWorkbook = true) replacement for the sequence
        ///   CreateNewWorkbook -> CreateWirelistReportHeader -> AppendToExcel -> SortWirelistWireNumber
        /// used by the sheet and component wire-list reports. Same header, same data block, same formatting, same
        /// Excel sort (so the same row order) - but the workbook is created with a single sheet, written completely
        /// in memory-to-Excel bulk calls and saved ONCE with SaveAs, instead of SaveAs + re-Open + several Saves +
        /// Close(true). Returns the .xls path, or "null" on failure like CreateNewWorkbook.
        /// </summary>
        public string WriteWirelistWorkbookOnePass(string iWorkbookName, string iSheetName, string reportType, object[,] iarr, int sortRowEnd)
        {
            try
            {
                ReportWB = ExcelApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
                ExcelApp.DisplayAlerts = false;

                ReportWS = ReportWB.Sheets[1];
                ReportWS.Name = iSheetName;

                ReportAppendRow = 1;
                WriteWirelistHeaderRows();

                AppendToExcel(iarr, 0, saveWorkbook: false);
                SortWirelistWireNumber(2, sortRowEnd, saveBeforeClose: false, closeWorkbook: false);

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
                }

                Logging.Info($"Wirelist Report (one pass) {workbookPath}.xls");
                return workbookPath + ".xls";
            }
            catch (COMException ex)
            {
                MessageBox.Show("Excel ##04: " + ex.Message);
                Logging.Info("Error interacting with Excel: " + ex.Message);
                return "null";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Excel ##05: " + ex.Message);
                Logging.Info("An error occurred: " + ex.Message);
                return "null";
            }
            finally
            {
                if (ReportWB != null)
                {
                    ReportWB.Close(false); // already saved by SaveAs above
                    Marshal.ReleaseComObject(ReportWB);
                }
            }
        }

        /// <summary>
        /// Closes the current report workbook WITHOUT saving. For reports that already saved it (AppendToExcel saves on
        /// success) and then left it open: those workbooks piled up in the Excel instance - one per loom - until Quit.
        /// If the save never happened the data was lost at Quit anyway, so nothing is lost by closing here.
        /// </summary>
        public void CloseReportWorkbookWithoutSaving()
        {
            try
            {
                if (ReportWB != null)
                {
                    ReportWB.Close(false);
                }
            }
            catch (Exception ex)
            {
                Logging.Error("CloseReportWorkbookWithoutSaving: " + ex.Message);
            }
        }

        // this method sort data an Excel sheet based on wire code for sheet,project wirelist, Component specifice breakdown reports
        public void SortWirelistWireNumber(int rowStart, int rowEnd, bool saveBeforeClose = true, bool closeWorkbook = true)
        {
            Excel.Range sortRange = null;
            Excel.Range keyRange = null;
            try
            {
                // The callers pass the height of the whole project array, but only the rows AppendToExcel actually wrote
                // (up to ReportAppendRow - 1) can hold data - everything below is empty and stays empty when sorted.
                // Sorting the full height was needless work and, on large jobs, went past the 65,536-row limit of the
                // .xls format: Excel rejected the range, the sort never ran and the report stayed unsorted.
                // Never reaches above the requested end, and never ends before the first data row.
                int lastWrittenRow = ReportAppendRow - 1;
                rowEnd = Math.Min(rowEnd, Math.Max(lastWrittenRow, rowStart));

                string sRange = $"A{rowStart}:J{rowEnd}";
                string colFilter1 = $"E{rowStart}:E{rowEnd}";

                // Define the range to sort
                sortRange = ReportWS.Range[sRange];
                keyRange = ReportWS.Range[colFilter1];

                // ReportWS.Sort is a COM property lookup every time it is written - fetch the object once
                var sort = ReportWS.Sort;

                // Clear any previous sort fields
                sort.SortFields.Clear();

                // Add the sorting field based on column E
                sort.SortFields.Add(
                    Key: keyRange,
                    SortOn: XlSortOn.xlSortOnValues,
                    Order: XlSortOrder.xlAscending
                // DataOption: XlSortDataOption.xlSortNormal
                );

                // Apply sorting properties and execute the sort
                sort.SetRange(sortRange);
                sort.Header = XlYesNoGuess.xlGuess;
                sort.MatchCase = false;
                sort.Orientation = XlSortOrientation.xlSortColumns;
                sort.SortMethod = XlSortMethod.xlPinYin;
                sort.Apply();
                // Save the workbook (the finally block below closes it with Close(true), which saves the same
                // content, so callers may skip this intermediate save)
                if (saveBeforeClose)
                {
                    ReportWB.Save();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Excel ##10: " + ex.Message);
                Logging.Error("An error occurred: " + ex.Message);
            }
            finally
            {
                if (closeWorkbook && ReportWB != null)
                {
                    ReportWB.Close(true);
                    Marshal.ReleaseComObject(ReportWB);
                }

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

        #region merging and Demerging

        public void SubLoomMergeCellsWithaValue(object iFrom, object iTo, object iVal1, object iVal2, object iNumOfShieldRemToMerge, object iCableClassification, Worksheet sheet = null)
        {
            // The caller has just activated this sheet; passing it avoids a late-bound ReportWB.ActiveSheet lookup per range.
            Worksheet ws = sheet ?? (Worksheet)ReportWB.ActiveSheet;
            // Merge cells in column B and set value

            Range rangeB = ws.Range["B" + iFrom, "B" + iTo];
            rangeB.Merge();
            rangeB.Value = iVal1;

            // Get value from the initial cell in column D and set for merged range
            Range rangeDInitial = ws.Range["D" + iFrom];
            var sLength = rangeDInitial.Value;
            Range rangeD = ws.Range["D" + iFrom, "D" + iTo];
            rangeD.Merge();
            rangeD.Value = sLength;

            // Merging cells in column K with conditional merging based on parameters           
            Range rangeK;
            if (Convert.ToInt32(iNumOfShieldRemToMerge) != 0 && iCableClassification == "Special")
            {
                rangeK = ws.Range["K" + iFrom, "K" + ((int)iTo + (int)iNumOfShieldRemToMerge)];
                rangeK.Merge();
                rangeK.Value = ""; // or iVal2 if required
            }
            else
            {
                rangeK = ws.Range["K" + iFrom, "K" + iTo];
                rangeK.Merge();
                rangeK.Value = iVal2;
            }
        }

        public void SubLoomUNmergeCellsWithaValue(object iFrom, object iTo, object iVal1, object iVal2, object iCableClassification, Worksheet sheet = null)
        {
            Worksheet ws = sheet ?? (Worksheet)ReportWB.ActiveSheet;
            // Unmerge cells in column B and set alignment properties
            Range rangeB = ws.Range["B" + iFrom, "B" + iTo];
            rangeB.MergeCells = false;
            rangeB.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            rangeB.VerticalAlignment = XlVAlign.xlVAlignCenter;
            rangeB.WrapText = true;
            rangeB.Orientation = 0;
            rangeB.AddIndent = false;
            rangeB.IndentLevel = 0;
            rangeB.ShrinkToFit = false;
            rangeB.ReadingOrder = -5002;//reading replace for XlReadingOrder.xlContext
            rangeB.Value = iVal1;

            // Conditional merge and value setting for column K based on Cable Classification
            Range rangeK = ws.Range["K" + iFrom, "K" + iTo];
            rangeK.Merge();

            if (iCableClassification == "Normal")
            {
                rangeK.Value = iVal2;
            }
            else if (iCableClassification == "Special")
            {
                rangeK.Value = ""; // Empty value for "Special"
            }
        }
        #endregion

        public void GenerateHALReportFormat_CableList(object[,] iarr, string reportName, int numberOfSheetsRequired, int loomSheetRowReq, string reportType)
        {
            string templatePath = Environment.GetEnvironmentVariable("ELECTRE_CUSTOMIZE") + @"\system\CABLE_ReportFormat.xls";

            try
            {
                ReportWB = ExcelApp.Workbooks.Add();
                ExcelApp.DisplayAlerts = false;

                // Open the template workbook
                TempWB = TemplateReadOnly ? ExcelApp.Workbooks.Open(templatePath, ReadOnly: true) : ExcelApp.Workbooks.Open(templatePath);

                string reportPath = Path.Combine(GlobalVar.ReportFolderGlobal, reportType);

                if (!Directory.Exists(reportPath))
                {
                    Directory.CreateDirectory(reportPath);
                }

                // .xls (BIFF8) holds at most XlsSheetLimit.MaxSheetsPerXls sheets - above that, Excel refuses the next
                // Copy() partway through the loop below and no file is produced. See XlsSheetLimit.
                var cableListFormat = XlsSheetLimit.ResolveFormat(reportName, numberOfSheetsRequired);
                string reportFilePath = Path.Combine(reportPath, cableListFormat.FileName);

                ReportWB.SaveAs(reportFilePath, cableListFormat.FileFormat);

                // The lookup below never changes across iterations (TempWB.Sheets is the same
                // small, fixed template collection every time), so find "CABLE LIST" once instead
                // of re-scanning it on every one of the numberOfSheetsRequired iterations. If not
                // found, the original would fail identically on its very first iteration (l=1) and
                // break before copying anything, so failing before the loop here is equivalent.
                Worksheet cableListSheet = null;
                foreach (Worksheet sheet in TempWB.Sheets)
                {
                    if (sheet.Name == "CABLE LIST")
                    {
                        cableListSheet = sheet;
                        break;
                    }
                }

                if (cableListSheet == null)
                {
                    Logging.Error("Sheet 'CABLE LIST' does not exist in the template workbook.");
                }
                else
                {
                    // Copy the required number of sheets (CL-1 .. CL-n) from the "CABLE LIST" template sheet
                    TemplateSheetCopier.CopyInto(ReportWB, cableListSheet, numberOfSheetsRequired, "CL-");
                }

                // Close the template workbook
                TempWB.Close(!TemplateReadOnly);
                Marshal.ReleaseComObject(TempWB);

                // Disable display alerts
                ExcelApp.DisplayAlerts = false;

                // Call DeleteDefaultSheets function, assumed to be defined elsewhere
                DeleteDefaultSheets();

                int Q = 1;

                // Populate each sheet with data from iarr
                for (int R = 1; R <= numberOfSheetsRequired; R++)
                {
                    Worksheet ws1 = (Worksheet)ReportWB.Sheets[R];
                    ws1.Activate();

                    if (R != 1)
                    {
                        //changes in first sheet will reflect in rest of the sheet, value are hard coded, to improve performace
                        ws1.Range["C48:C52"].Formula = "='CL-1'!C48";   // relative: C49 -> 'CL-1'!C49 ... same as five single-cell writes
                        
                        ws1.Range["D48"].NumberFormat = "General";
                        ws1.Range["D48:D52"].Formula = "='CL-1'!D48";
                        ws1.Range["F51"].Value = ws1.Parent.Worksheets["CL-1"].Range["F51"].Text;

                        ws1.Range["C48:C52"].NumberFormat = "@";
                        ws1.Range["D48:D52"].NumberFormat = "dd-mm-yyyy";
                        ws1.Columns["D"].AutoFit();

                        
                        ws1.Range["F49"].Formula = "='CL-1'!F49";
                        ws1.Range["F51"].Formula = "='CL-1'!F51";
                        ws1.Range["F52"].Formula = "='CL-1'!F52";
                        ws1.Range["K51:K52"].Formula = "='CL-1'!K51";
                    }

                    LoomInitialRow = 8;
                    LoomInitialCol = 2;

                    // ---- Pass 1: collect this sheet's Q-values (identical break-before-increment bookkeeping as
                    // before), then write the WHOLE sheet's raw row data in ONE bulk call instead of one call per
                    // row (was up to loomSheetRowReq separate calls per sheet). The strikethrough/merge/unmerge
                    // pass below is deferred to run AFTER all of this sheet's raw data is written, including for
                    // "ME" groups whose iTo row is written later in this same block. This reordering was verified
                    // directly in Excel (COM) against the original interleaved order for both the "ME" and "UN"
                    // cases: identical final cell values and identical merge geometry in both orders - only the
                    // call count differs. (Not just reasoned about - actually run and diffed in real Excel.)
                    var sheetQs = new List<int>(loomSheetRowReq);
                    for (int O = 1; O <= loomSheetRowReq; O++)
                    {
                        sheetQs.Add(Q);
                        if (Q < iarr.GetLength(0) - 1)
                        {
                            Q++;
                        }
                        else
                        {
                            break;
                        }
                    }

                    if (sheetQs.Count > 0)
                    {
                        // Column 10 (index 9) intentionally ends up as iarr[Q,19], not iarr[Q,10] -
                        // the original code wrote iarr[Q,10] via the P-loop and then immediately
                        // overwrote that same cell with iarr[Q,19]; only the final value is kept here.
                        object[,] sheetBuffer = new object[sheetQs.Count, 10];
                        for (int r = 0; r < sheetQs.Count; r++)
                        {
                            int rowQ = sheetQs[r];
                            for (int P = 1; P <= 9; P++)
                            {
                                sheetBuffer[r, P - 1] = iarr[rowQ, P];
                            }
                            sheetBuffer[r, 9] = iarr[rowQ, 19];
                        }

                        ExcelBulk.WriteValues(ws1, LoomInitialRow + 1, LoomInitialCol + 1, sheetBuffer);

                        // ---- Pass 2: strikethrough / merge / unmerge / last, in the same Q order as before.
                        for (int r = 0; r < sheetQs.Count; r++)
                        {
                            int rowQ = sheetQs[r];
                            int excelRow = LoomInitialRow + 1 + r;

                            if (iarr[rowQ, 19].ToString() == "D")
                            {

                                //  Range headerRange = ReportWS.Range["A1", "I1"];
                                Range strikeThroughRange = ws1.Range["B" + excelRow, "L" + excelRow];
                                if (strikeThroughRange != null)
                                {
                                    strikeThroughRange.MergeCells = false;  // Ensure cells are not merged
                                    strikeThroughRange.Font.Strikethrough = true; // Apply strikethrough
                                    // (Removed the per-row "ScreenUpdating = true" and ws1.Calculate(): Excel is hidden and
                                    // Calculation is Automatic, so neither changes the saved file, but the first one switched
                                    // the screen-updating optimisation back on for the rest of the run.)
                                }
                            }

                            if (!string.IsNullOrEmpty((string)iarr[rowQ, 12]))
                            {
                                if (iarr[rowQ, 11].ToString() == "ME")
                                {
                                    SubLoomMergeCellsWithaValue(iarr[rowQ, 12], iarr[rowQ, 13], iarr[rowQ, 14], iarr[rowQ, 15], iarr[rowQ, 16], iarr[rowQ, 18], ws1);
                                }
                                else if (iarr[rowQ, 11].ToString() == "UN")
                                {
                                    SubLoomUNmergeCellsWithaValue(iarr[rowQ, 12], iarr[rowQ, 12], iarr[rowQ, 14], iarr[rowQ, 15], iarr[rowQ, 18], ws1);
                                }
                                else if (iarr[rowQ, 11].ToString() == "LAST")
                                {
                                    ws1.Cells[excelRow, LoomInitialCol].Value = iarr[rowQ, 14];
                                }
                            }
                        }
                    }

                    ws1.Range["L51:L52"].FormulaR1C1 = $"SHEET {R} OF {ReportWB.Sheets.Count} SHEETS";
                    ws1.Range["B3:D4"].FormulaR1C1 = "";
                    //adding loomName and issue 
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
                    ws1.Range["B47"].Value = DateTime.Now.Year;
                    alignCellsXl(8, 46, ws1);
                }
               
                ReportWB.Save();
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
                // Clean up resources              
                if (ReportWB != null)
                {
                    ReportWB.Close(true);
                    Marshal.ReleaseComObject(ReportWB);
                }
            }
        }
        
        public void alignCellsXl(int fromCell, int toCell, Worksheet sheet = null)
        {
            Worksheet ws = sheet ?? (Worksheet)ReportWB.ActiveSheet;
            Range rangeB = ws.Range["B" + fromCell, "O" + toCell];
            rangeB.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            rangeB.VerticalAlignment = XlVAlign.xlVAlignCenter;
        }
        // this method generates Continuity List report format with HAL template
        public void GenerateHALReportFormat_ContinuityList(object[,] arrFTcwob, string reportType, string reportName)
        {
            int numRowInReportSheet = 47;
            int StartRow;
            int StartCol;
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

            try
            {
                ReportWB = ExcelApp.Workbooks.Add();
                ExcelApp.DisplayAlerts = false;

                TempWB = ExcelApp.Workbooks.Open(templatePath);

                string reportPath = Path.Combine(GlobalVar.ReportFolderGlobal, reportType);

                if (!Directory.Exists(reportPath))
                {
                    Directory.CreateDirectory(reportPath);
                }

                // Same XlsSheetLimit reasoning as GenerateHALReportFormat_CableList.
                var continuityFormat = XlsSheetLimit.ResolveFormat(reportName, numberOfSheetsRequired);
                string reportFilePath = Path.Combine(reportPath, continuityFormat.FileName);

                ReportWB.SaveAs(reportFilePath, continuityFormat.FileFormat);

                // Same hoist as GenerateHALReportFormat_CableList: the lookup never changes across
                // iterations, so find it once instead of on every one of the numberOfSheetsRequired
                // iterations.
                Worksheet cableListSheet = null;
                foreach (Worksheet sheet in TempWB.Sheets)
                {
                    if (string.Equals(sheet.Name, "continuity", StringComparison.OrdinalIgnoreCase))
                    {
                        cableListSheet = sheet;
                        break;
                    }
                }

                if (cableListSheet == null)
                {
                    Logging.Error("Sheet 'CABLE LIST' does not exist in the template workbook.");
                }
                else
                {
                    // Copy the required number of sheets (CWOB-1 .. CWOB-n) from the "continuity" template sheet
                    TemplateSheetCopier.CopyInto(ReportWB, cableListSheet, numberOfSheetsRequired, "CWOB-");
                }

                // Close the template workbook
                TempWB.Close(true);
                Marshal.ReleaseComObject(TempWB);

                // Disable display alerts
                ExcelApp.DisplayAlerts = false;

                // Call DeleteDefaultSheets function, assumed to be defined elsewhere
                DeleteDefaultSheets();

                int Q = 0;
                for (int R = 1; R <= numberOfSheetsRequired; R++)
                {
                    Worksheet ws1 = (Worksheet)ReportWB.Sheets[R];
                    ws1.Activate();
                    ws1.Cells[4, 12].Value = $"{R} OF {numberOfSheetsRequired}";

                    if (R != 1)
                    {
                        //changes in first sheet will reflect in rest of the sheet, value are hard coded, to improve performace
                        ws1.Range["C56:C57"].Formula = "='CWOB-1'!C56";
                        ws1.Range["F56"].Formula = "='CWOB-1'!F56";
                        ws1.Range["I56"].Formula = "='CWOB-1'!I56";
                        ws1.Range["L56"].Formula = "='CWOB-1'!L56";
                        ws1.Range["F4"].Formula = "='CWOB-1'!F4";
                        ws1.Range["I4"].Formula = "='CWOB-1'!I4";
                    }

                    StartRow = 6;
                    StartCol = 3;

                    // Collect this sheet's rows first (same Q-advance/exhaustion logic as before),
                    // then issue one bulk Range.Value write per destination column instead of one
                    // COM call per cell. The gap columns (StartCol+1,4,7,9,10) are static template
                    // content the original loop never touched, so they're still left alone.
                    void WriteColumnBulk(int col, List<object> values)
                    {
                        if (values.Count == 0) return;
                        object[,] buffer = new object[values.Count, 1];
                        for (int r = 0; r < values.Count; r++) buffer[r, 0] = values[r];
                        ExcelBulk.WriteValues(ws1, StartRow, col, buffer);
                    }

                    var col0 = new List<object>();
                    var col2 = new List<object>();
                    var col3 = new List<object>();
                    var col5 = new List<object>();
                    var col6 = new List<object>();
                    var col8 = new List<object>();
                    var col11 = new List<object>();
                    bool exhausted = false;

                    for (int O = StartRow; O <= StartRow + numRowInReportSheet; O++)
                    {
                        col0.Add(arrFTcwob[Q, 0]);
                        col2.Add(arrFTcwob[Q, 1]);
                        col3.Add(arrFTcwob[Q, 2]);
                        col5.Add(arrFTcwob[Q, 3]);
                        col6.Add(arrFTcwob[Q, 4]);
                        col8.Add(arrFTcwob[Q, 5]);
                        col11.Add(arrFTcwob[Q, 6]);

                        if (Q < arrFTcwob.GetLength(0) - 1)
                        {
                            Q++;
                        }
                        else
                        {
                            exhausted = true;
                            break;
                        }
                    }

                    WriteColumnBulk(StartCol, col0);
                    WriteColumnBulk(StartCol + 2, col2);
                    WriteColumnBulk(StartCol + 3, col3);
                    WriteColumnBulk(StartCol + 5, col5);
                    WriteColumnBulk(StartCol + 6, col6);
                    WriteColumnBulk(StartCol + 8, col8);
                    WriteColumnBulk(StartCol + 11, col11);

                    if (exhausted)
                    {
                        ReportWB.Save();
                    }

                    // alignCellsXl(6,53) is idempotent (fixed range, fixed properties, no dependency
                    // on loop state), so calling it once has the same final effect as the original
                    // per-iteration calls - except when the loop broke on its very first iteration,
                    // in which case the original never reached that call at all for this sheet.
                    bool calledAlignOriginally = !exhausted || col0.Count > 1;
                    if (calledAlignOriginally)
                    {
                        alignCellsXl(6, 53, ws1);
                    }
                }
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
                // Clean up resources              

                if (ReportWB != null)
                {
                    ReportWB.Close(true);
                    Marshal.ReleaseComObject(ReportWB);
                }
            }
        }
        // after generating Excel reports this method releseComObjactes
        public void releaseExcel()
        {
            ExcelApp.Quit();
            Marshal.ReleaseComObject(ExcelApp);
        }

      // this method create header for sheet,project wirelist,component specifice breakdown reports 
        public bool CreateWirelistReportHeader(string workbookPath)
        {
            ReportAppendRow = 1;

            try
            {
                ReportWB = ExcelApp.Workbooks.Open(workbookPath);

                ReportWS = ReportWB.ActiveSheet;

                if (ReportWS == null)
                {
                    Logging.Error("Failed to access the active sheet.");
                    return false;
                }

                WriteWirelistHeaderRows();

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

        public string AppendToExcelPowerOn(List<ElectreObject> elecCollection)
        {
            List<string> parameterPINs = new List<string>();
            try
            {
                var sourceComponents = ComponentSelectors.ScbTcbSource(elecCollection);
                if (sourceComponents.Count == 0)
                {
                    MessageBox.Show("Circuit breakers not found in the Panel", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Logging.Info("Circuit breakers not found in the Panel");
                    return "null";
                }

                // Ground candidates grouped by panel once, instead of a full scan of elecCollection per circuit breaker.
                var groundIndex = new GroundCandidateIndex(elecCollection);

                foreach (var source in sourceComponents)
                {
                    var visitedConnections = new HashSet<string>();

                    //Get the Negative/Ground pins for the source component
                    List<string> negPins = PowerOnReport.FindGroundPin(groundIndex, source);

                    // Tracing the Destination component from Source
                    TraceAndLogPath(source, source.WireNumber, source.ConnectorName, source.PinNumber, source.SubNet, visitedConnections, negPins, ref parameterPINs);
                }
                if (parameterPINs.Count >= 2)
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
                ReportWS.Columns[1].AutoFit();
                ReportWS.Columns[5].AutoFit();
                ReportWB.Save();
                return ReportWS.Name;
            }
            catch (Exception ex)
            {
                Logging.Error("AppendToExcelPowerOn: Error while appending PowerOn data: " + ex.Message);
                return "null";
            }
        }

        // Helper Method: Recursive path tracing
        private void TraceAndLogPath(ElectreObject source, string wireNumber, string connectorName, string pinNumber, string subNet, HashSet<string> visited, List<string> negPins,ref List<string> parameterPINs)
        {
            try
            {
                visited.Add($"{connectorName},{pinNumber}");

                // Reuses PowerOnReport's shared modMain.ElecCollection_All index (the same one its
                // Trace* helpers already use below) instead of rescanning the whole project-wide
                // collection on every recursive call.
                var connectedObjects = PowerOnReport.GetConnectedObjectsByWireSubNet(wireNumber, subNet)
                    .Where(w => !visited.Contains($"{w.ConnectorName},{w.PinNumber}"))
                    .ToList();

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
                            var nextObjEQU = ConnectorPinTracer.TracePinofEQUConnector(PowerOnReport.GetIndex(), connObj);
                            if (nextObjEQU == null)
                            {
                                LogToExcel(source, connObj.ConnectorName, connObj.PinNumber, negPins, ref parameterPINs);
                                return;
                            }
                            if (nextObjEQU != null)
                                TraceAndLogPath(source, nextObjEQU.WireNumber, nextObjEQU.ConnectorName, nextObjEQU.PinNumber, nextObjEQU.SubNet, visited, negPins, ref parameterPINs);
                            break;
                        case "DIS":
                            var nextObj = ConnectorPinTracer.TracePinofBreakConnector(PowerOnReport.GetIndex(), connObj);
                            if (nextObj == null)
                            {
                                LogToExcel(source, connObj.ConnectorName, connObj.PinNumber, negPins, ref parameterPINs);
                                return;
                            }
                            if (nextObj != null)
                                TraceAndLogPath(source, nextObj.WireNumber, nextObj.ConnectorName, nextObj.PinNumber, nextObj.SubNet, visited, negPins, ref parameterPINs);
                            break;

                        case "TBK":
                            var jmList = ConnectorPinTracer.TracePinOfJM(PowerOnReport.GetIndex(), connObj);
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
                            var splList = ConnectorPinTracer.TracePinOfSPL(PowerOnReport.GetIndex(), connObj);
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
                            var terList = ConnectorPinTracer.TracePinOfTER(PowerOnReport.GetIndex(), connObj);
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
                // Build the source row plus one row per (comma-containing) negative pin first,
                // then issue a single bulk Range.Value write instead of one COM call per cell -
                // this is called once per traced connection, so it's the hottest write path in
                // Power-On report generation.
                var rowConnector = new List<string>();
                var rowPin = new List<string>();

                if (negPins.Count == 0)
                {
                    rowConnector.Add("wrt STR");
                    rowPin.Add("-");
                }
                else
                {
                    foreach (var negPin in negPins)
                    {
                        if (negPin.Contains(","))
                        {
                            string groundConnector = negPin.Split(',')[0];
                            string groundPin = negPin.Split(',')[1];
                            rowConnector.Add(groundConnector);
                            rowPin.Add($"{groundPin}(-)");
                        }
                    }
                }

                int totalRows = rowConnector.Count + 1;

                object[,] buffer = new object[totalRows, 5];
                buffer[0, 0] = source.Panel;
                buffer[0, 1] = source.ConnectorName;
                buffer[0, 2] = destConnector;
                buffer[0, 3] = $"{destPin}(+)";
                buffer[0, 4] = source.Voltage;

                for (int r = 0; r < rowConnector.Count; r++)
                {
                    buffer[r + 1, 1] = "";
                    buffer[r + 1, 2] = rowConnector[r];
                    buffer[r + 1, 3] = rowPin[r];
                }

                ExcelBulk.WriteValues(ReportWS, ReportAppendRow, 1, buffer);

                ExcelBulk.RangeAt(ReportWS, ReportAppendRow, 5, ReportAppendRow + totalRows - 1, 5).Merge();
                ExcelBulk.RangeAt(ReportWS, ReportAppendRow, 2, ReportAppendRow + totalRows - 1, 2).Merge();

                if (string.Equals(source.ComponentType, "SCB", StringComparison.OrdinalIgnoreCase))
                {
                    ExcelBulk.RangeAt(ReportWS, ReportAppendRow, 1, ReportAppendRow + totalRows - 1, 1).Merge();
                }
                else if (string.Equals(source.ComponentType, "TCB", StringComparison.OrdinalIgnoreCase))
                {
                    // ReportWS.Range[ReportWS.Cells[ReportAppendRow, 1], ReportWS.Cells[ReportAppendRow + 3, 1]].Merge();
                    ExcelBulk.RangeAt(ReportWS, ReportAppendRow, 1, ReportAppendRow + totalRows + 1, 1).Merge();
                }

                ReportWS.Rows[ReportAppendRow].AutoFit();
                ReportWS.Rows[ReportAppendRow + totalRows - 1].AutoFit();
                // ReportAppendRow += 3;
                ReportAppendRow += totalRows + 1;

                if (string.Equals(source.ComponentType, "TCB", StringComparison.OrdinalIgnoreCase))
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

        #region AppendToExcelPowerOn old methods
        /*public string AppendToExcelPowerOn(List<ElectreObject> elecCollection)
        {
            try
            {
                // Step 1: Get only SCBs & TCBs as source components
                var sourceComponents = ComponentSelectors.ScbTcbSource(elecCollection);
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
                    if (source.ComponentType.ToUpper() == "SCB")
                    {
                        ReportWS.Range[ReportWS.Cells[ReportAppendRow, 1], ReportWS.Cells[ReportAppendRow + 1, 1]].Merge();
                    }
                    else if(source.ComponentType.ToUpper() == "TCB")
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

                    if (source.ComponentType.ToUpper() == "TCB")
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
                 var sourceComponents = ComponentSelectors.ScbTcbSource(elecCollection);
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
                     if (source.ComponentType.ToUpper() == "SCB")
                     {
                         ReportWS.Range[ReportWS.Cells[ReportAppendRow, 1], ReportWS.Cells[ReportAppendRow + 1, 1]].Merge();
                     }
                     else if (source.ComponentType.ToUpper() == "TCB")
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

                     if (source.ComponentType.ToUpper() == "TCB")
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
            var sourceComponents = ComponentSelectors.ScbTcbSource(elecCollection);
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

                    if (source.ComponentType.ToUpper() == "SCB")
                    {
                        ReportWS.Range[ReportWS.Cells[ReportAppendRow, 1], ReportWS.Cells[ReportAppendRow + 1, 1]].Merge();
                    }
                    else if (source.ComponentType.ToUpper() == "TCB")
                    {
                        ReportWS.Range[ReportWS.Cells[ReportAppendRow, 1], ReportWS.Cells[ReportAppendRow + 3, 1]].Merge();
                    }

                    ReportAppendRow += 3;

                    if (source.ComponentType.ToUpper() == "TCB")
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
                ReportWB = ExcelApp.Workbooks.Open(workbookPath);

                ReportWS = ReportWB.ActiveSheet as Worksheet;

                if (ReportWS == null)
                {
                    Logging.Error("Failed to access the active sheet.");
                    return false;
                }

                // Set header cells
                ReportWS.Cells[ReportAppendRow, 1].Value = "Distribution Box";
                ReportWS.Cells[ReportAppendRow, 2].Value = "CB to be pressed";
                ReportWS.Cells[ReportAppendRow, 3].Value = "Unit";
                ReportWS.Cells[ReportAppendRow, 4].Value = "Pins";
                ReportWS.Cells[ReportAppendRow, 5].Value = "Parameter";
                ReportWS.Cells[ReportAppendRow, 6].Value = "Remarks (MeasuredVoltage)";

                //  ReportWS.Cells[ReportAppendRow, "Shunt"].Value = "Shunt"; 
                // Uncomment if Shunt is required

                // Format the header row
                Range headerRange = ReportWS.Range["A1", "F1"];
                headerRange.Font.Bold = true;

                // Center the text in the header row
                headerRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;

                // Set column widths
                ReportWS.Columns["A"].ColumnWidth = 20;
                ReportWS.Columns["B"].ColumnWidth = 20;
                ReportWS.Columns["C"].ColumnWidth = 15;
                ReportWS.Columns["D"].ColumnWidth = 15;
                ReportWS.Columns["E"].ColumnWidth = 20;
                ReportWS.Columns["F"].ColumnWidth = 30;

                // Center the text in all columns for data rows as well
                Range dataRange = ReportWS.Range["A2", "F" + ReportWS.Rows.Count];
                dataRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                dataRange.VerticalAlignment = XlHAlign.xlHAlignCenter;

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
       
        private void DeleteDefaultSheets(Excel.Workbook wb)
        {
            foreach (Excel.Worksheet sheet in wb.Sheets)
            {
                // Only delete if not your copied sheets (optional condition)
                if (sheet.Name != "MS" && !sheet.Name.StartsWith("ML-"))
                {
                    sheet.Delete();
                }
            }
        }

        public bool ReadOOTBMaterialListXL()
        {
            var projectPath = Path.GetDirectoryName(GlobalVar.ReportFolderGlobal.TrimEnd('\\')) + "\\";
            string projectName = new DirectoryInfo(projectPath.TrimEnd('\\')).Name;
            try
            {
                string ootbWBms_xls = projectPath + "result\\" + projectName + "_mat.xls";
                string ootbWBms_xlsx = projectPath + "result\\" + projectName + "_mat.xlsx";
                Excel.Workbook ootbWBms = null;

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
                // Excel.Worksheet OOTBmsSheet = ootbWBms.Sheets[modMain.ChkListPanelText] as Excel.Worksheet;
                Excel.Worksheet OOTBmsSheet = ootbWBms.Sheets["F"] as Excel.Worksheet;

                if (OOTBmsSheet == null)
                {
                    Console.WriteLine("The Panel MS sheet is not in " + ootbWBms.Name);
                    Console.WriteLine("Panel MS is not generated.");
                    return false;
                }

                int i = 1;
                while (OOTBmsSheet.Cells[i, 1].Value != null)
                {
                    for (int S = 1; S <= 9; S++)
                    {
                        // modMain.arrTableOfOOTBms_Panel[i, S] = OOTBmsSheet.Cells[i, S].Value;
                        var cellValue = OOTBmsSheet.Cells[i, S].Value;
                        modMain.arrTableOfOOTBms_Panel[i, S] = cellValue != null ? cellValue.ToString() : string.Empty;
                    }
                    i++;
                }
                modMain.OOTBmsTotalrow_Selection = i;

                // Read weight from "WEIGHT RESULT" sheet
                OOTBmsSheet = ootbWBms.Sheets["WEIGHT RESULT"] as Excel.Worksheet;
                for (int yy = 1; yy <= 200; yy++)
                {
                    var panelName = OOTBmsSheet.Cells[yy, 1].Value;
                    if (panelName != null && panelName.ToString() == modMain.ChkListPanelText)
                    {
                        OOTBPaneltotalWeight = OOTBmsSheet.Cells[yy, 2].Value;
                        break;
                    }
                }

                ootbWBms.Close(false);
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
                return false;
            }
        }

        public void GenerateHALReportFormat_MaterialList(string[] iarr, string iReportName)
        {
            Excel.Application excelApp = null;
            Excel.Workbook TemplateWB = null;
            Excel.Workbook ReportWB = null;
            try
            {
                if (!ReadOOTBMaterialListXL()) return; // Check if the OOTB material list is created

                int NumRowInReportSheet = 16; // Number of rows per sheet in the report
                int StartRow = 13; // Start Row for data
                int StartCol = 2; // Start Column for data
                int NumberOfSheetsRequired = (int)Math.Floor((double)modMain.OOTBmsTotalrow_Selection / NumRowInReportSheet + 1); // Calculate the number of sheets needed
                bool bReportTemplate = false;

                bReportTemplate = File.Exists(Environment.GetEnvironmentVariable("ELECTRE_CUSTOMIZE") + "\\system\\MATERIALLIST_ReportFormat.xls");
                if (!bReportTemplate)
                {
                    Console.WriteLine("The MaterialList Report Template is missing. So saving file into TEMPFILES folder without template");
                    return;
                }
                string panelDrawingPath = ConfigurationManager.AppSettings["PanelDrawingFolder"];
                string templFolder = Path.Combine(GlobalVar.ReportFolderGlobal, panelDrawingPath);

                /*Excel.Application excelApp = new Excel.Application();
                Excel.Workbook TemplateWB = excelApp.Workbooks.Open(Environment.GetEnvironmentVariable("ELECTRE_CUSTOMIZE") + "\\system\\MATERIALLIST_ReportFormat.xls");
                Excel.Workbook ReportWB = excelApp.Workbooks.Add();
               // ReportWB.SaveAs(frmMain.lblReportFolder.Caption + iReportName, Excel.XlFileFormat.xlExcel8);
                ReportWB.SaveAs(templFolder + iReportName, Excel.XlFileFormat.xlExcel8);


                for (int l = 1; l <= NumberOfSheetsRequired; l++)
                {
                   *//* excelApp.ActiveWindow.ActivateNext();
                    Excel.Sheets sheets = ReportWB.Sheets;
                    Excel.Worksheet sheetMS = (Excel.Worksheet)sheets["MS"];
                    sheetMS.Copy(After: ReportWB.Sheets[ReportWB.Sheets.Count]);
                    ReportWB.Sheets[ReportWB.Sheets.Count].Name = "ML-" + l;*//*

                    Excel.Worksheet sheetMS = (Excel.Worksheet)TemplateWB.Sheets["MS"];
                    sheetMS.Copy(After: ReportWB.Sheets[ReportWB.Sheets.Count]);
                    ReportWB.Sheets[ReportWB.Sheets.Count].Name = "ML-" + l;
                }

                TemplateWB.Close();

                excelApp.DisplayAlerts = false;
                DeleteDefaultSheets();*/

                excelApp = new Excel.Application();
                excelApp.DisplayAlerts = false; // Set early to suppress popups

                string templatePath = Environment.GetEnvironmentVariable("ELECTRE_CUSTOMIZE") + "\\system\\MATERIALLIST_ReportFormat.xls";

                if (!File.Exists(templatePath))
                {
                    Console.WriteLine("Template file not found.");
                    return;
                }

                TemplateWB = excelApp.Workbooks.Open(templatePath);
                ReportWB = excelApp.Workbooks.Add();

                // Copy "MS" sheet from template to report
                for (int l = 1; l <= NumberOfSheetsRequired; l++)
                {
                    Excel.Worksheet templateSheet = null;
                    try
                    {
                        templateSheet = (Excel.Worksheet)TemplateWB.Sheets["MS"];
                    }
                    catch
                    {
                        Console.WriteLine("Template sheet 'MS' not found.");
                        TemplateWB.Close(false);
                        ReportWB.Close(false);
                        excelApp.Quit();
                        return;
                    }

                    templateSheet.Copy(After: ReportWB.Sheets[ReportWB.Sheets.Count]);
                    ReportWB.Sheets[ReportWB.Sheets.Count].Name = "ML-" + l;
                }

                DeleteDefaultSheets(ReportWB); // Clean up default sheets

                // Save the workbook initially
                ReportWB.SaveAs(Path.Combine(templFolder, iReportName), Excel.XlFileFormat.xlExcel8);

                TemplateWB.Close(false);
     /*           ReportWB.Save(); // Save final result

                // Clean up Excel instance
                ReportWB.Close(true);
                excelApp.Quit();*/

                int Q = 1;

                // Initialize array for material details (like part number, weight, etc.)
                string[,] arrDetails = new string[500, 10];
                int L1, L2;

                for (int S = 1; S <= NumberOfSheetsRequired; S++)
                {
                    Excel.Worksheet WS1 = (Excel.Worksheet)ReportWB.Sheets[S];
                    WS1.Activate();
                    WS1.Cells[9, 2] = $"SHEET {S} OF {NumberOfSheetsRequired} SHEETS";
                    // WS1.Cells[8, 2] = frmMain.txtPanelDigit.Text;
                    WS1.Cells[8, 2] = PanelDrawingWindow.panelDigit;
                    WS1.Cells[9, 8] = OOTBPaneltotalWeight;

                    for (int R = StartRow; R < StartRow + NumRowInReportSheet; R++)
                    {
                        WS1.Cells[R, StartCol] = Q; // serial number
                        WS1.Cells[R, StartCol + 2] = modMain.arrTableOfOOTBms_Panel[Q + 1, 4]; // Quantity
                        WS1.Cells[R, StartCol + 3] = modMain.arrTableOfOOTBms_Panel[Q + 1, 5]; // Part Number
                        WS1.Cells[R, StartCol + 5] = modMain.arrTableOfOOTBms_Panel[Q + 1, 7]; // Description
                        WS1.Cells[R, StartCol + 6] = modMain.arrTableOfOOTBms_Panel[Q + 1, 9]; // Weight

                        string ak1 = modMain.arrTableOfOOTBms_Panel[Q + 1, 2];
                        if (ak1.StartsWith("LOC-"))
                        {
                            WS1.Cells[R, StartCol + 9] = ak1.Substring(4); // Remove LOC-
                        }
                        else
                        {
                            WS1.Cells[R, StartCol + 9] = ak1; // Ref Connector
                        }

                        if (Q < modMain.OOTBmsTotalrow_Selection - 2)
                        {
                            Q++;
                        }
                        else
                        {
                            ReportWB.Save();
                            return; // Exit after saving
                        }
                    }
                }

                ReportWB.Save();
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
            finally
            {
                try
                {
                    if (ReportWB != null)
                    {
                        ReportWB.Close(true);
                        Marshal.ReleaseComObject(ReportWB);
                    }

                    if (TemplateWB != null)
                    {
                        TemplateWB.Close(false);
                        Marshal.ReleaseComObject(TemplateWB);
                    }

                    if (excelApp != null)
                    {
                        excelApp.Quit();
                        Marshal.ReleaseComObject(excelApp);
                    }
                }
                catch (Exception cleanupEx)
                {
                    Console.WriteLine("Error during Excel cleanup: " + cleanupEx.Message);
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
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
            TextFileAppender.AppendLines(filePath, materialsList);
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
                List<ElectreObject> panelDetails = modMain.ElecCollection_All
                .Where(e => e.Panel == ChkListPanelText)
                .ToList();
                int missingDestCount = 0;
                foreach (var obj in panelDetails)
                {   
                    var sourceCon = obj.ConnectorName;
                    var sourcePin = obj.PinNumber;
                    // Reuses PowerOnReport's shared modMain.ElecCollection_All index (narrows by
                    // WireNumber+SubNet in O(1)) instead of rescanning the full project-wide
                    // collection for every panelDetails row.
                    var destinationObj = PowerOnReport.GetConnectedObjectsByWireSubNet(obj.WireNumber, obj.SubNet)
                                                              //.Where(e => e.Core_Part_Number == obj.Core_Part_Number)
                                                              .FirstOrDefault(e => e.ConnectorName != obj.ConnectorName);
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
            TextFileAppender.AppendLines(filePath, materialsList);
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

