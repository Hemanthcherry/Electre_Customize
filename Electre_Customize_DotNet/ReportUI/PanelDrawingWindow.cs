using Electre_Customize_DotNet.Contracts;
using Electre_Customize_DotNet.MainOperation;
using Electre_Customize_DotNet.Objects;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Application = System.Windows.Forms.Application;
using CheckBox = System.Windows.Forms.CheckBox;
using Label = System.Windows.Forms.Label;
using ListBox = System.Windows.Forms.ListBox;
using Excel = Microsoft.Office.Interop.Excel;
using Electre_Customize_DotNet.Logs;

namespace Electre_Customize_DotNet.ReportUI
{
    public partial class PanelDrawingWindow : UserControl
    {
        private HashSet<string> checkedPanelItems;
        private ISelectionFunctions _selectionFunctions;
        private bool Status;
        private HashSet<string> checkedLoomItems;
        private HashSet<string> checkedSheetItems;
        private List<ElectreObject> selectedPanelList = new List<ElectreObject>();
        public static string panelDigit = string.Empty;

        private static object lockObjet = new object();

        private static PanelDrawingWindow? _panelWindowInst;
        public static PanelDrawingWindow PanelWindowInst
        {
            get
            {
                if (_panelWindowInst == null)
                {
                    lock (lockObjet)
                    {
                        if (_panelWindowInst == null)
                        {
                            _panelWindowInst = new PanelDrawingWindow();
                        }
                    }
                }

                return _panelWindowInst;
            }
        }


        public List<string> SelectedPanelListPanelDrawing
        {
            get
            {
                return checkedPanelItems.ToList();
            }
        }

        //public string ChkListPanelText
        //{
        //    get { return chkListPanel.Text; }
        //    set { chkListPanel.Text = value; }
        //}

        public string ChkListPanelText
        {
            get
            {
                if(chkListPanel.InvokeRequired)
                {
                    return (string)chkListPanel.Invoke(new Func<string>(() => chkListPanel.Text));
                }
                else
                {
                    return chkListPanel.Text;
                }
            }
            set
            {
                if (chkListPanel.InvokeRequired)
                {
                    chkListPanel.Invoke(new System.Action(() => chkListPanel.Text = value));
                }
                else
                {
                    chkListPanel.Text = value;
                }
            }
        }
        private PanelDrawingWindow()
        {
            InitializeComponent();
            checkedLoomItems = new HashSet<string>();//keep the selected Items.
            checkedSheetItems = new HashSet<string>();
            checkedPanelItems = new HashSet<string>();
            _selectionFunctions = SelectionFunctions.SelectionFuntionsInstance;
            var filteredItems = modMain.arrListOfPANEL
                .Where(item => !string.IsNullOrWhiteSpace(item))
                .ToList();
            _selectionFunctions.UpdateCheckedListBox(filteredItems, checkedPanelItems, chkListPanel);
            modMain.arrListOfLOOM = modMain.arrListOfLOOM.OrderBy(x => x).ToList();
            modMain.arrListOfSHEET = modMain.arrListOfSHEET.OrderBy(x => x).ToList();
        }

        // commented on september 10th, 2025 before implementing one selection at a time
        /* private void chkListPanel_ItemCheck(object sender, ItemCheckEventArgs e)
         {
             _selectionFunctions.ListItemChecked(e, checkedPanelItems, chkListPanel, lblPanelCount);
         }*/

        private void chkListPanel_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            // Prevent recursive calls
            chkListPanel.ItemCheck -= chkListPanel_ItemCheck;

            if (e.NewValue == CheckState.Checked)
            {
                // Uncheck all other items except the one being checked
                for (int i = 0; i < chkListPanel.Items.Count; i++)
                {
                    if (i != e.Index && chkListPanel.GetItemChecked(i))
                    {
                        chkListPanel.SetItemChecked(i, false);
                        checkedPanelItems.Remove(chkListPanel.Items[i].ToString());
                    }
                }
                // Add the currently checked item
                checkedPanelItems.Add(chkListPanel.Items[e.Index].ToString());
            }
            else
            {
                // Remove the item if unchecked
                checkedPanelItems.Remove(chkListPanel.Items[e.Index].ToString());
            }

            // Update the label count
            // string searchCount = lblPanelCount.Text;
            // lblPanelCount.Text = "Selected Items Count: " + checkedPanelItems.Count;

            // Reattach event handler
            chkListPanel.ItemCheck += chkListPanel_ItemCheck;
        }


        // Sample stub for the FindAndSelectInList method (implement as needed)
        private void FindAndSelectInList(string connectorName, object listControl)
        {
            // This method should implement the logic for finding and selecting the connector
            // in the specified list control (listEQU, listJUN, etc.)
            ListBox listBox = listControl as ListBox;
            if (listBox != null)
            {
                for (int i = 0; i < listBox.Items.Count; i++)
                {
                    if (listBox.Items[i].ToString() == connectorName)
                    {
                        listBox.SelectedIndex = i;
                        break;
                    }
                }
            }
        }

        private void ClearAndDisableCheckbox(CheckBox iCheck)
        {
            iCheck.Checked = false;
            iCheck.Enabled = false;
        }

        private void ClearAndEnableCheckbox(CheckBox iCheck)
        {
            iCheck.Checked = false;
            iCheck.Enabled = true;
        }

        private bool ValidateFolderSelection(string fileLocation)
        {
            return File.Exists(fileLocation);
        }
        public bool SelectAllInList(ListBox iList)
        {
            if (iList.Items.Count == 0)
            {
                MessageBox.Show("No item in list " + iList.Name);
                return false;
            }

            for (int i = 0; i < iList.Items.Count; i++)
            {
                iList.SetSelected(i, true);
            }

            return true;
        }

        // Deselect All Items in the ListBox
        public bool DeSelectAllInList(ListBox iList)
        {
            if (iList.Items.Count == 0)
            {
                MessageBox.Show("No item in list " + iList.Name);
                return false;
            }

            for (int i = 0; i < iList.Items.Count; i++)
            {
                iList.SetSelected(i, false);
            }

            return true;
        }

        // Add Array Items to ListBox and Sort it
        public void AddArrayItemToList(List<string> arr, ListBox iListbox)
        {
            foreach (var item in arr)
            {
                iListbox.Items.Add(item);
            }

            SortListBoxLogical(iListbox);
        }

        // Sort the ListBox
        private void SortListBoxLogical(ListBox LB)
        {
            var items = LB.Items.Cast<string>().ToList();
            items.Sort(StringComparer.OrdinalIgnoreCase);

            LB.Items.Clear();
            foreach (var item in items)
            {
                LB.Items.Add(item);
            }
        }

        // Check if Folder Exists
        public bool FolderExists(string sFullPath)
        {
            return Directory.Exists(sFullPath);
        }

        private void FindAndSelectInList(string searchStr, ListBox listBox)
        {
            for (int b = 0; b < listBox.Items.Count; b++)
            {
                if (string.Equals(searchStr, listBox.Items[b].ToString(), StringComparison.OrdinalIgnoreCase))
                {
                    listBox.SetSelected(b, true);
                    break;
                }
            }
        }
        private void FindAndSelectMultiple(List<string> searchList, ListBox listBox)
        {
            listBox.ClearSelected();

            for (int i = 0; i < listBox.Items.Count; i++)
            {
                string currentItem = listBox.Items[i].ToString();
                if (searchList.Any(s => string.Equals(s, currentItem, StringComparison.OrdinalIgnoreCase)))
                {
                    listBox.SetSelected(i, true);
                }
            }
        }

        public void RunExcelAutomation(string templFolder)
        {
            Excel.Application xlApp = null;
            Excel.Workbook xlBook = null;

            try
            {
                // string tempFolder = frmMain.lblTempFolder.Text;
                string customizePath = Environment.GetEnvironmentVariable("ELECTRE_CUSTOMIZE");
                string userName = Environment.GetEnvironmentVariable("USERNAME");

                xlApp = new Excel.Application();
                xlApp.Visible = false;

                // Open the correct workbook
                string defaultWorkbook = System.IO.Path.Combine(customizePath, @"system\WirelistScript_Vexec.xlsm");
                string altWorkbook = @"E:\work_dsi\Customers\HAL\Electre_Customization\WirelistScript_Vexec.xlsm";

                xlBook = xlApp.Workbooks.Open(userName == "K1O" ? altWorkbook : defaultWorkbook, ReadOnly: true);

                // Run main macro with initial parameter
                xlApp.Run("MainMod", templFolder);

                xlApp.DisplayAlerts = false;
                xlApp.Quit();

                // Re-launch for further operations
                xlApp = new Excel.Application();
                xlBook = xlApp.Workbooks.Open(defaultWorkbook, ReadOnly: true);

                // Run DelimitInput macro
                string csvPath = System.IO.Path.Combine(templFolder, "Mod_CATData.csv");
                xlApp.Run("DelimitInput", csvPath);

                xlApp.DisplayAlerts = false;
                xlApp.Quit();
            }
            catch (COMException ex)
            {
                MessageBox.Show("There was an automation error.\nQuitting...\n" + ex.Message, "Automation Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Unexpected error:\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Release COM objects to avoid memory leaks
                if (xlBook != null) Marshal.ReleaseComObject(xlBook);
                if (xlApp != null) Marshal.ReleaseComObject(xlApp);

                xlBook = null;
                xlApp = null;

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private void btnProceed_Click(object sender, EventArgs e)
        {
            string panelDrawingFolderPath = ConfigurationManager.AppSettings["PanelDrawingFolder"];
            string templFolder = Path.Combine(GlobalVar.ReportFolderGlobal, panelDrawingFolderPath);
            /* if (chkListPanel.SelectedItems.Count == 0) 
             {
                 var res = MessageBox.Show("Please select the Panel!!", "Information");
                 return;
             }*/
            if (checkedPanelItems.Count == 0)
            {
                var res = MessageBox.Show("Please select the Panel!!", "Information");
                return;
            }

            /*  var result = MessageBox.Show("Is the OOTB Material list created?", "Confirmation", MessageBoxButtons.YesNo);
              if (result != DialogResult.Yes)
              {
                  MessageBox.Show("Please create the ELECTRE MaterialList before proceeding with Panel Drawing.\nExiting...", "Critical", MessageBoxButtons.OK, MessageBoxIcon.Error);
                  Application.Exit();
              }*/

            List<string> selectedPanels = checkedPanelItems.ToList();

            panelDigit = Microsoft.VisualBasic.Interaction.InputBox("Enter the 15 digit number for panel", "Input");
            txtPanelDigit.Text = panelDigit;

            // selectedPanelList = modMain.ElecCollection_All.Where(e => e.Panel == selectedPanel).ToList();
            selectedPanelList = modMain.ElecCollection_All.Where(e => selectedPanels.Contains(e.Panel)).ToList();

            List<string> panelLoomList = selectedPanelList.Select(e => e.BundleName).ToList();
            List<string> panelSheetList = selectedPanelList.Select(e => e.SheetName).ToList();
            List<string> panelList = selectedPanelList.Select(e => e.Panel).ToList();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            string wordDocumentPath = ConfigurationManager.AppSettings["panelExtractionDocumentPath"];

            if (!string.IsNullOrEmpty(wordDocumentPath))
            {
                OpenWordDocument(wordDocumentPath);
            }
            else
            {
                MessageBox.Show($"Document not found in the following path : {wordDocumentPath}");
                Logging.Warning($"Document not found in the following path : {wordDocumentPath}");
            }
        }

        private void OpenWordDocument(string filePath)
        {
            try
            {
                var psi = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = filePath,
                    UseShellExecute = true
                };
                System.Diagnostics.Process.Start(psi);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error opening document: {ex.Message}");
            }
        }
    }
}
