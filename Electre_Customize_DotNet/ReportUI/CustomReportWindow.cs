using Electre_Customize_DotNet.Exceptions;
using Electre_Customize_DotNet.Logs;
using Electre_Customize_DotNet.MainOperation;
using Electre_Customize_DotNet.Objects;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Diagnostics.Contracts;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Electre_Customize_DotNet.ReportUI
{
    public partial class CustomReportWindow : Form
    {
        private bool isMainFormOpen = false;

        //Creating Objects for Each User Controls
        private CableListWindow _cableListWindow;
        private ComponentBreak _componentBreak;
        private ContinuityWindow _continuityWindow;
        private MeggerScheduler _meggerScheduler;
        private PowerOnWindow _powerOnWindow;
        public PanelDrawingWindow _panelDrawingWindow;  // changed public from private for accessing the ChkListPanelText variable in modMain
        private PanelDrawingSchedulesWindow _panelDrawingSchedulesWindow;

        private Button _activeMenuButton = null;
        private readonly Color MenuDefaultBackColor = Color.CornflowerBlue;
        private readonly Color MenuDefaultForeColor = Color.White;
        private readonly Color MenuSelectedBackColor = Color.Orange;

        private modMain _modMain;
        public List<string> SelectedLoomList = new List<string>();
        public List<string> SelectedSheetList = new List<string>();
        public bool projectWireList;
        public List<string> SelectedEqupList = new List<string>();
        public List<string> SelectedBrkList = new List<string>();
        public List<string> SelectedJmList = new List<string>();
        public List<string> SelectedMiscList = new List<string>();
        public List<string> SelectedContLoomList = new List<string>();
        public List<string> SelectedMeggerLoomList = new List<string>();
        public bool meggersheet;
        public bool powerOnProjectList;
        public List<string> SelectedPanelList = new List<string>();
        public List<string> SelectedPanelListPanelDrawing = new List<string>();

        public List<string> SelectedSheetListPanelDwg = new List<string>();
        public List<string> SelectedLoomListPanelDwgCL = new List<string>();
        public List<string> SelectedSheetListPanelDwgCL = new List<string>();
        public List<string> SelectedLoomListPanelDwgCont = new List<string>();
        public List<string> SelectedLoomListPanelDwgMeg = new List<string>();
        public List<string> SelectedSheetListPanelDwgPower = new List<string>();


        public CustomReportWindow()
        {
            InitializeComponent();
        }

        private void LoadUserControl(UserControl userControl, string title)
        {
            panelMain.Controls.Clear();

            //Clearing lists before generating the report for the selected items.
            SelectedSheetList.Clear();
            SelectedLoomList.Clear();
            SelectedMiscList.Clear();
            SelectedJmList.Clear();
            SelectedBrkList.Clear();
            SelectedEqupList.Clear();
            SelectedContLoomList.Clear();
            SelectedMeggerLoomList.Clear();
            SelectedPanelListPanelDrawing.Clear();

            SelectedLoomListPanelDwgCL.Clear();
            SelectedSheetListPanelDwgCL.Clear();
            SelectedLoomListPanelDwgCont.Clear();
            SelectedLoomListPanelDwgMeg.Clear();
            SelectedSheetListPanelDwgPower.Clear();

            //lblTitle.Text = title;
            panelMain.Controls.Add(userControl);

            // If the user control is cableListWindow, save the reference
            if (userControl is CableListWindow)
            {
                _cableListWindow = (CableListWindow)userControl;
            }

            // If the user control is componentBreakDown, save the reference
            if (userControl is ComponentBreak)
            {
                _componentBreak = (ComponentBreak)userControl;
            }

            // If the user control is ContinuityWindow, save the reference
            if (userControl is ContinuityWindow)
            {
                _continuityWindow = (ContinuityWindow)userControl;
            }
            if (userControl is MeggerScheduler)
            {
                _meggerScheduler = (MeggerScheduler)userControl;
            }
            // If the user control is PowerOnWindow, save the reference
            if (userControl is PowerOnWindow)
            {
                _powerOnWindow = (PowerOnWindow)userControl;
            }
            // If the user control is PanelDrawing, save the reference
            if (userControl is PanelDrawingWindow)
            {
                _panelDrawingWindow = (PanelDrawingWindow)userControl;
            }
            if (userControl is PanelDrawingSchedulesWindow)
            {
                _panelDrawingSchedulesWindow = (PanelDrawingSchedulesWindow)userControl;
            }
        }

        private void CreateHiddenTempFolder(string tempFolderPath)
        {
            if (!Directory.Exists(tempFolderPath))
            {
                DirectoryInfo di = Directory.CreateDirectory(tempFolderPath);
                di.Attributes |= FileAttributes.Hidden;
            }
        }

        private void CustomReportWindow_Load(object sender, EventArgs e)
        {
            try
            {

                if (!isMainFormOpen)
                {
                    /* DialogResult dialogResult = MessageBox.Show(
                       "Do you want to proceed with the ECM process?",
                        "ECM Process",
                      MessageBoxButtons.YesNo,
                       MessageBoxIcon.Question);*/

                    CustomMessageBox.CustomResult dialogResult;
                    using (var msgBox = new CustomMessageBox("Do you want to proceed with the ECM process?", "ECM Process"))
                    {
                        if (msgBox.ShowDialog() == DialogResult.OK)
                        {
                            dialogResult = msgBox.Result;
                        }
                        else
                        {
                            // User closed the dialog or canceled
                            Application.Exit();
                            return;
                        }
                    }

                    //if (dialogResult == DialogResult.Yes)
                    if (dialogResult == CustomMessageBox.CustomResult.WithConfig)
                    {
                        using (var variantsForm = new ECM_VariantsList())
                        {
                            var dialogResultVar = variantsForm.ShowDialog();

                            if (dialogResultVar == DialogResult.OK)
                            {
                                //var selectedVariants = variantsForm.chkListVarients.CheckedItems.Cast<string>().ToList();
                                // Get the selected variant name
                                string selectedVariant = ECM_VariantsList.SelectedVariant;
                                //MessageBox.Show($"Selected Variant '{selectedVariant}'");

                                // Then proceed with your ECM process
                                if (!Directory.Exists(GlobalVar.TempFolderGlobal))
                                {
                                    CreateHiddenTempFolder(GlobalVar.TempFolderGlobal);
                                }
                                variantName.Text = $"With Configuration:\n{selectedVariant}";
                                RunDataExtractionBatchECM(GlobalVar.StrtCmd, selectedVariant);

                                if (File.Exists(GlobalVar.DataExtractionGlobal))
                                {
                                    modMain.MainFunction();
                                    modMain.MainFunction_withPanels(); // added for Power On
                                }
                                else
                                {
                                    MessageBox.Show("Please check for the below\n1. data_extraction file is missing\n2. Data Extraction of this project is not done");
                                    Application.Exit();
                                    // return;
                                }
                            }
                            else
                            {
                                // User cancelled or closed the popup form
                                // MessageBox.Show("ECM variants selection cancelled.");
                                Application.Exit();
                                // return;
                            }
                        }
                        /*   // Proceed with ECM process
                           if (!Directory.Exists(GlobalVar.TempFolderGlobal))
                           {
                               CreateHiddenTempFolder(GlobalVar.TempFolderGlobal);
                           }

                           RunDataExtractionBatchECM(GlobalVar.StrtCmd,"schem_ecm_copy");

                           if (File.Exists(GlobalVar.DataExtractionGlobal))
                           {
                               modMain.MainFunction();
                               modMain.MainFunction_withPanels(); // added for Power On
                           }
                           else
                           {
                               MessageBox.Show("Please check for the below\n1. data_extraction file is missing\n2. Data Extraction of this project is not done");
                               Application.Exit();
                               return;
                           }*/
                    }
                    else
                    {
                        if (!Directory.Exists(GlobalVar.TempFolderGlobal))
                        {
                            CreateHiddenTempFolder(GlobalVar.TempFolderGlobal);
                        }


                        RunDataExtractionBatch(GlobalVar.StrtCmd);


                        if (File.Exists(GlobalVar.DataExtractionGlobal))
                        {
                            modMain.MainFunction();
                            modMain.MainFunction_withPanels();
                        }
                        else
                        {
                            MessageBox.Show("Please check for the below\n1. data_extraction file is missing\n2. Data Extraction of this project is not done");
                            Application.Exit();
                            return;
                        }
                    }
                    isMainFormOpen = true;
                }

                Cursor = Cursors.Default;
            }

            catch (DuplicatWireException ex)
            {
                Logging.Error(ex.Message);
                this.Close();
            }

            catch (Exception ex)
            {
                Logging.Error(ex.Message);
                this.Close();
            }

        }

        private void RunDataExtractionBatch(string projectPath)
        {
            if (string.IsNullOrEmpty(projectPath))
            {
                MessageBox.Show("Project Empty for DataExtraction");
                return;
            }

            // Set paths based on input arguments and environment variables
            string prjPath = projectPath;
            string rootDir = Environment.GetEnvironmentVariable("ROOTDIR") ?? string.Empty;
            string electreCustomize = Environment.GetEnvironmentVariable("ELECTRE_CUSTOMIZE") ?? string.Empty;

            if (string.IsNullOrEmpty(rootDir) || string.IsNullOrEmpty(electreCustomize))
            {
                Logging.Error("Environment variables ROOTDIR or ELECTRE_CUSTOMIZE are not set.");
                return;
            }

            //initializing the dataExtraction path, dxs file, .d files

            string instPath = Path.Combine(rootDir, "system", "bin");
            string dxsPath = Path.Combine(electreCustomize, "system", "dxs");

            string dataExtractionFilePath = Path.Combine(prjPath, "schema", "data_extraction.csv");
            string dxsFilePath = Path.Combine(dxsPath, "my_levelNET_csv.dxs");

            string schemPath = string.Empty;

            schemPath = Path.Combine(prjPath, "schem", "*");


            // Run the new_data_extract.exe command
            var processInfo = new ProcessStartInfo
            {
                FileName = Path.Combine(instPath, "new_data_extract.exe"),
                Arguments = ":E" + " " + $"{dataExtractionFilePath}" + " " + $"{dxsFilePath}" + " " + $"{schemPath}",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            //creating the data_extraction.csv file and writng the data

            using (var process = new Process { StartInfo = processInfo })
            {
                process.Start();
                string output = process.StandardOutput.ReadToEnd();
                string error = process.StandardError.ReadToEnd();
                process.WaitForExit();

                if (!string.IsNullOrEmpty(output))
                {
                    MessageBox.Show("Output: " + output);
                }

                if (!string.IsNullOrEmpty(error))
                {
                    MessageBox.Show("Error: " + error);
                }
            }
        }

        private void btnCable_Click(object sender, EventArgs e)
        {
            SetActiveMenu(btnCable);
            LoadUserControl(CableListWindow.CableListWindowInst, ConfigurationManager.AppSettings["cableListTitle"]);
        }

        private void btnCompBrk_Click(object sender, EventArgs e)
        {
            SetActiveMenu(btnCompBrk);
            LoadUserControl(ComponentBreak.ComponentBreakInst, ConfigurationManager.AppSettings["componetBreakTitle"]);
        }

        private void btnCont_Click(object sender, EventArgs e)
        {
            SetActiveMenu(btnCont);
            LoadUserControl(ContinuityWindow.ContinuityWindowInst, ConfigurationManager.AppSettings["continuityTitle"]);
        }
        private void btnPowerOn_Click(object sender, EventArgs e)
        {
            SetActiveMenu(btnPowerOn);
            LoadUserControl(PowerOnWindow.PowerOnWindowInst, ConfigurationManager.AppSettings["powerOnTitle"]);
        }

        private void btnMegger_Click(object sender, EventArgs e)
        {
            SetActiveMenu(btnMegger);
            LoadUserControl(MeggerScheduler.MeggerSchedulerinst, ConfigurationManager.AppSettings["MeggerTitle"]);
        }

        private void btnPanel_Click(object sender, EventArgs e)
        {
            SetActiveMenu(btnPanel);
            LoadUserControl(PanelDrawingWindow.PanelWindowInst, ConfigurationManager.AppSettings["panelDrawingTitle"]);
        }

        private void btnPanelSchedules_Click(object sender, EventArgs e)
        {
            SetActiveMenu(btnpanelSchedules);
            LoadUserControl(PanelDrawingSchedulesWindow.PanelSchedulesWindowInst, ConfigurationManager.AppSettings["panelDwgSchedulesTitle"]);
        }

        private void SetActiveMenu(Button selectedButton)
        {
            ResetAllMenuButtons();

            selectedButton.BackColor = MenuSelectedBackColor;
            selectedButton.ForeColor = MenuDefaultForeColor;
           // selectedButton.FlatAppearance.BorderColor = Color.DarkOrange;

            _activeMenuButton = selectedButton;
        }

        private void ResetAllMenuButtons()
        {
            foreach (Control ctrl in panelLeft.Controls)
            {
                if (ctrl is Button btn)
                {
                    btn.BackColor = MenuDefaultBackColor;
                    btn.ForeColor = MenuDefaultForeColor;
                }
            }
        }

        private async void btnOk_Click(object sender, EventArgs e)
        {
            if (_cableListWindow != null)
            {
                projectWireList = _cableListWindow.ProjWirelist;
                SelectedLoomList = _cableListWindow.SelectedLoomList;
                SelectedSheetList = _cableListWindow.SelectedSheetList;
            }

            if (_componentBreak != null)
            {
                SelectedEqupList = _componentBreak.SelectedEqupList;
                SelectedBrkList = _componentBreak.SelectedBrkList;
                SelectedJmList = _componentBreak.SelectedJmList;
                SelectedMiscList = _componentBreak.SelectedMiscList;
            }

            if (_continuityWindow != null)
            {
                SelectedContLoomList = _continuityWindow.SelectedLoomList;
            }
            if (_meggerScheduler != null)
            {
                // meggersheet = _meggerScheduler.Meggerscheduler;
                SelectedMeggerLoomList = _meggerScheduler.SelectedLoomList;
            }
            if (_powerOnWindow != null)
            {
                powerOnProjectList = _powerOnWindow.ProjWirelist;
                SelectedPanelList = _powerOnWindow.SelectedPanelList;
            }
            if (_panelDrawingWindow != null)
            {
                // powerOnProjectList = _powerOnWindow.ProjWirelist;
                SelectedPanelListPanelDrawing = _panelDrawingWindow.SelectedPanelListPanelDrawing;
            }

            if (_panelDrawingSchedulesWindow != null)
            {
                SelectedLoomListPanelDwgCL = _panelDrawingSchedulesWindow.SelectedLoomListCL;
                SelectedSheetListPanelDwgCL = _panelDrawingSchedulesWindow.SelectedSheetListCL;
                SelectedLoomListPanelDwgCont = _panelDrawingSchedulesWindow.SelectedLoomListCon;
                SelectedLoomListPanelDwgMeg = _panelDrawingSchedulesWindow.SelectedLoomListMeg;
                SelectedSheetListPanelDwgPower = _panelDrawingSchedulesWindow.SelectedSheetListPower;

            }

            LoadingForm loadingForm = new LoadingForm();
            _modMain = new modMain(this, loadingForm);

            loadingForm.Show();
            Application.DoEvents();

            await Task.Run(() => { _modMain.ReportExcecution(); });
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void btnOk_MouseHover(object sender, EventArgs e)
        {
            showToolTip(btnOk, "Generate Report");
            btnOk.BackColor = Color.Green;
            btnOk.ForeColor = Color.Blue;
        }

        private void btnClose_MouseHover(object sender, EventArgs e)
        {
            showToolTip(btnClose, "Close Application");
            btnClose.BackColor = Color.Red;
            btnClose.ForeColor = Color.Blue;
        }

        private void btnCable_MouseHover(object sender, EventArgs e)
        {
            showToolTip(btnCable, "Cable List");
        }

        private void showToolTip(Button btn, string info)
        {
            InfoTtp.SetToolTip(btn, info);
        }

        private void btnCompBrk_MouseHover(object sender, EventArgs e)
        {
            showToolTip(btnCompBrk, "Component BreakDown");
        }

        private void btnCont_MouseHover(object sender, EventArgs e)
        {
            showToolTip(btnCont, "Continuity W/O BreakDown");
        }

        private void btnMegger_MouseHover(object sender, EventArgs e)
        {
            showToolTip(btnMegger, "Megger Schedules");
        }

        private void btnPowerOn_MouseHover(object sender, EventArgs e)
        {
            showToolTip(btnPowerOn, "Power On Schedules");
        }

        private void btnPanelDwg_MouseHover(object sender, EventArgs e)
        {
            showToolTip(btnPanel, "Data Extraction for Panel Drawing");
        }

        private void btnPanelDwgSchedules_MouseHover(object sender, EventArgs e)
        {
            showToolTip(btnpanelSchedules, "Panel Drawing Schedules");
        }

        private void btnOk_MouseLeave(object sender, EventArgs e)
        {
            changeToDefaultColor(btnOk);
        }

        private void changeToDefaultColor(Button btn)
        {
            btn.BackColor = Color.Lavender;
            btn.ForeColor = Color.Black;
        }

        private void btnClose_MouseLeave(object sender, EventArgs e)
        {
            changeToDefaultColor(btnClose);
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            if (_cableListWindow != null)
            {
                // ClearAllCheckboxes(_cableListWindow);
                ClearAll(_cableListWindow);
            }

            if (_componentBreak != null)
            {
                // ClearAllCheckboxes(_componentBreak);
                ClearAll(_componentBreak);
            }

            if (_continuityWindow != null)
            {
                //ClearAllCheckboxes(_continuityWindow);
                ClearAll(_continuityWindow);
            }

            if (_powerOnWindow != null)
            {
                ClearAll(_powerOnWindow);
            }

            if (_meggerScheduler != null)
            {
                ClearAll(_meggerScheduler);
            }

            if (_panelDrawingWindow != null)
            {
                ClearAll(_panelDrawingWindow);
            }

            if (_panelDrawingSchedulesWindow != null)
            {
                ClearAll(_panelDrawingSchedulesWindow);
            }

        }

        // feb 26 1:45pm

        private void ClearAll(Control parentControl)
        {
            ClearAllTextBoxes(parentControl);
            ClearAllCheckboxes(parentControl);
        }

        private void ClearAllTextBoxes(Control parentcontrol)
        {
            foreach (Control control in parentcontrol.Controls)
            {
                if (control is TextBox textbox)
                {
                    textbox.Text = "Search Filter";
                    //textbox.Clear();
                }
                else if (control.HasChildren)
                {
                    ClearAllTextBoxes(control);
                }
            }
        }

        private void ClearAllCheckboxes(Control parentControl)
        {
            // var selectionFunctions = SelectionFunctions.SelectionFuntionsInstance;
            // selectionFunctions.SuppressListItemEvents = true;

            foreach (Control control in parentControl.Controls)
            {
                if (control is CheckBox checkBox)
                {
                    checkBox.Checked = false; // Uncheck CheckBox
                }
                else if (control is CheckedListBox checkedListBox)
                {
                    checkedListBox.BeginUpdate();
                    for (int i = 0; i < checkedListBox.Items.Count; i++)
                    {
                        checkedListBox.SetItemChecked(i, false);
                    }
                    checkedListBox.ClearSelected(); // removes blue highlight
                    checkedListBox.EndUpdate();
                }
                else if (control.HasChildren)
                {
                    ClearAllCheckboxes(control); // Recurse into children
                }
            }

            // selectionFunctions.SuppressListItemEvents = false;
        }

        // Commented on July 10th Before implementing flickering Issue code in UI
        /*  private void ClearAllCheckboxes(Control parentControl)
          {
              foreach (Control control in parentControl.Controls)
              {
                  if (control is CheckBox checkBox)
                  {
                      checkBox.Checked = false; // Uncheck CheckBox
                  }
                  else if (control is CheckedListBox checkedListBox)
                  {
                      for (int i = 0; i < checkedListBox.Items.Count; i++)
                      {
                          checkedListBox.SetItemChecked(i, false); // Uncheck CheckedListBox items
                      }
                  }
                  else if (control.HasChildren)
                  {
                      // Recursively call for child controls
                      ClearAllCheckboxes(control);
                  }
              }
          }*/


        private void RunDataExtractionBatchECM(string projectPath, string selectedVariant)
        {
            if (string.IsNullOrEmpty(projectPath))
            {
                MessageBox.Show("Project Empty for DataExtraction");
                return;
            }

            // Set paths based on input arguments and environment variables
            string prjPath = projectPath;
            string rootDir = Environment.GetEnvironmentVariable("ROOTDIR") ?? string.Empty;
            string electreCustomize = Environment.GetEnvironmentVariable("ELECTRE_CUSTOMIZE") ?? string.Empty;

            if (string.IsNullOrEmpty(rootDir) || string.IsNullOrEmpty(electreCustomize))
            {
                Logging.Error("Environment variables ROOTDIR or ELECTRE_CUSTOMIZE are not set.");
                return;
            }

            //initializing the dataExtraction path, dxs file, .d files

            string instPath = Path.Combine(rootDir, "system", "bin");
            string dxsPath = Path.Combine(electreCustomize, "system", "dxs");

            string dataExtractionFilePath = Path.Combine(prjPath, "schema", "data_extraction.csv");
            string dxsFilePath = Path.Combine(dxsPath, "my_levelNET_csv.dxs");
            string ecmschemPath = Path.Combine(prjPath, selectedVariant, "*");
            string ecmfolderpath = Path.Combine(prjPath, selectedVariant);
            /*  string ecmschemPath = Path.Combine(prjPath, "schem_ecm_copy", "*");
              string ecmfolderpath = Path.Combine(prjPath, "schem_ecm_copy");*/
            if (!Directory.Exists(ecmfolderpath))
            {
                MessageBox.Show($"The folder {selectedVariant} does not exist. Exiting application.");
                Environment.Exit(0);  // Exit the application
            }


            // Run the new_data_extract.exe command
            var processInfo = new ProcessStartInfo
            {
                FileName = Path.Combine(instPath, "new_data_extract.exe"),
                Arguments = ":E" + " " + $"{dataExtractionFilePath}" + " " + $"{dxsFilePath}" + " " + $"{ecmschemPath}",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            //creating the data_extraction.csv file and writng the data

            using (var process = new Process { StartInfo = processInfo })
            {
                process.Start();
                string output = process.StandardOutput.ReadToEnd();
                string error = process.StandardError.ReadToEnd();
                process.WaitForExit();

                if (!string.IsNullOrEmpty(output))
                {
                    MessageBox.Show("Output: " + output);
                }

                if (!string.IsNullOrEmpty(error))
                {
                    MessageBox.Show("Error: " + error);
                }
            }
        }

    }
}
