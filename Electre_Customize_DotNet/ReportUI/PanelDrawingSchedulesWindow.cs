using Electre_Customize_DotNet.Contracts;
using Electre_Customize_DotNet.MainOperation;
using Electre_Customize_DotNet.Objects;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Electre_Customize_DotNet.ReportUI
{
    public partial class PanelDrawingSchedulesWindow : UserControl
    {
        private HashSet<string> checkedPanelItems;
        private ISelectionFunctions _selectionFunctions;
        private bool Status;
        private HashSet<string> checkedLoomItemsCL;
        private HashSet<string> checkedLoomItemsCon;
        private HashSet<string> checkedLoomItemsMeg;
        private HashSet<string> checkedSheetItems;
        private HashSet<string> checkedSheetItemsCL;
        private HashSet<string> checkedSheetItemsPower;

        private List<ElectreObject> selectedPanelList = new List<ElectreObject>();
        public static string panelDigit = string.Empty;

        public static List<ElectreObject> filteredSheetCollection = new List<ElectreObject>();

        private static object lockObjet = new object();

        private static PanelDrawingSchedulesWindow? _panelSchedulesWindowInst;
        public static PanelDrawingSchedulesWindow PanelSchedulesWindowInst
        {
            get
            {
                if (_panelSchedulesWindowInst == null)
                {
                    lock (lockObjet)
                    {
                        if (_panelSchedulesWindowInst == null)
                        {
                            _panelSchedulesWindowInst = new PanelDrawingSchedulesWindow();
                        }
                    }
                }

                return _panelSchedulesWindowInst;
            }
        }

        public List<string> SelectedLoomListCL
        {
            get
            {
                return checkedLoomItemsCL.ToList();
            }
        }

        public List<string> SelectedSheetListCL
        {
            get
            {
                return checkedSheetItemsCL.ToList();
            }
        }

        public List<string> SelectedLoomListCon
        {
            get
            {
                return checkedLoomItemsCon.ToList();
            }
        }

        public List<string> SelectedLoomListMeg
        {
            get
            {
                return checkedLoomItemsMeg.ToList();
            }
        }

        public List<string> SelectedSheetListPower
        {
            get
            {
                return checkedSheetItemsPower.ToList();
            }
        }

        public PanelDrawingSchedulesWindow()
        {
            InitializeComponent();

            modMain.arrListOfLOOM = modMain.arrListOfLOOM.OrderBy(x => x).ToList();
            modMain.arrListOfSHEET = modMain.arrListOfSHEET.OrderBy(x => x).ToList();
            filteredSheetCollection = modMain.ElecCollection;

            checkedSheetItems = new HashSet<string>();
            checkedLoomItemsCL = new HashSet<string>();
            checkedSheetItemsCL = new HashSet<string>();
            checkedLoomItemsCon = new HashSet<string>();
            checkedLoomItemsMeg = new HashSet<string>();
            checkedSheetItemsPower = new HashSet<string>();

            _selectionFunctions = SelectionFunctions.SelectionFuntionsInstance;
            // modMain.arrListOfLOOM = modMain.arrListOfLOOM.OrderBy(x => x).ToList();
            // modMain.arrListOfSHEET = modMain.arrListOfSHEET.OrderBy(x => x).ToList();
            _selectionFunctions.UpdateCheckedListBox(modMain.arrListOfSHEET, checkedSheetItems, chkListSheets);
            _selectionFunctions.UpdateCheckedListBox(modMain.arrListOfLOOM, checkedLoomItemsCL, chklistLoomCL);
            // _selectionFunctions.TotalCount(totalloom);
            _selectionFunctions.UpdateCheckedListBox(modMain.arrListOfSHEET, checkedSheetItemsCL, chkListSheetCL);

            _selectionFunctions.UpdateCheckedListBox(modMain.arrListOfLOOM, checkedLoomItemsCon, chkListLoomCont);
            _selectionFunctions.UpdateCheckedListBox(modMain.arrListOfLOOM, checkedLoomItemsMeg, chkListLoomMeg);
            _selectionFunctions.UpdateCheckedListBox(modMain.arrListOfSHEET, checkedSheetItemsPower, chkListSheetPowerOn);
        }

        private void chkSheet_CheckedChanged(object sender, EventArgs e)
        {
            if (chkSheet.CheckState != CheckState.Indeterminate)
            {
                _selectionFunctions.SelectCheckItems(chkSheet, chkListSheets, checkedSheetItems);
            }
        }

        private void chkListSheet_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            _selectionFunctions.ListItemCheckedPanelDwg(e, checkedSheetItems, chkListSheets, chkSheet);
        }

        private void chkLoomCL_CheckedChanged(object sender, EventArgs e)
        {
            if (chkLoomCL.CheckState != CheckState.Indeterminate)
            {
                _selectionFunctions.SelectCheckItems(chkLoomCL, chklistLoomCL, checkedLoomItemsCL);
            }
        }

        private void chkListLoomCL_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            _selectionFunctions.ListItemCheckedPanelDwg(e, checkedLoomItemsCL, chklistLoomCL, chkLoomCL);
        }

        private void chkSheetCL_CheckedChanged(object sender, EventArgs e)
        {
            if (chkSheetCL.CheckState != CheckState.Indeterminate)
            {
                _selectionFunctions.SelectCheckItems(chkSheetCL, chkListSheetCL, checkedSheetItemsCL);
            }
        }

        private void chkListSheetCL_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            _selectionFunctions.ListItemCheckedPanelDwg(e, checkedSheetItems, chkListSheetCL, chkSheetCL);
        }

        private void chkLoomCon_CheckedChanged(object sender, EventArgs e)
        {
            if (chkLoomCon.CheckState != CheckState.Indeterminate)
            {
                _selectionFunctions.SelectCheckItems(chkLoomCon, chkListLoomCont, checkedLoomItemsCon);
            }
        }

        private void chkListLoomCon_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            _selectionFunctions.ListItemCheckedPanelDwg(e, checkedLoomItemsCon, chkListLoomCont, chkLoomCon);
        }

        private void chkLoomMeg_CheckedChanged(object sender, EventArgs e)
        {
            if (chkLoomMeg.CheckState != CheckState.Indeterminate)
            {
                _selectionFunctions.SelectCheckItems(chkLoomMeg, chkListLoomMeg, checkedLoomItemsMeg);
            }
        }

        private void chkListLoomMeg_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            _selectionFunctions.ListItemCheckedPanelDwg(e, checkedLoomItemsMeg, chkListLoomMeg, chkLoomMeg);
        }

        private void chkSheetPowerOn_CheckedChanged(object sender, EventArgs e)
        {
            if (chkSheetPower.CheckState != CheckState.Indeterminate)
            {
                _selectionFunctions.SelectCheckItems(chkSheetPower, chkListSheetPowerOn, checkedSheetItemsPower);
            }
        }

        private void chkListSheetPowerOn_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            _selectionFunctions.ListItemCheckedPanelDwg(e, checkedSheetItemsPower, chkListSheetPowerOn, chkSheetPower);
        }

        private void btnFilter_Click(object sender, EventArgs e)
        {
            // Step 1: Get checked sheet names from the UI
            List<string> checkedSheetItems = chkListSheets.CheckedItems.Cast<string>().ToList();

            // Step 2: If nothing is selected, show message and exit
            if (checkedSheetItems.Count == 0)
            {
                MessageBox.Show("Please select at least one sheet to filter.",
                                "No Selection",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Warning);
                return;
            }

            // Step 3: Filter ElecCollection based on selected sheets
            filteredSheetCollection = modMain.ElecCollection
                .Where(e => checkedSheetItems.Contains(e.SheetName))
                .ToList(); // Store filtered data separately

            // Step 4: Update arrListOfSHEET with distinct, sorted SheetNames from the filtered data
            modMain.arrListOfSHEET = filteredSheetCollection
                .Select(e => e.SheetName)
                .Where(name => !string.IsNullOrEmpty(name))
                .Distinct()
                .OrderBy(name => name)
                .ToList();

            // Step 5: Update arrListOfLOOM based on the same filtered data (change logic if needed)
            modMain.arrListOfLOOM = filteredSheetCollection
                .Select(e => e.BundleName) // or use SheetName again if that's intended
                .Where(name => !string.IsNullOrEmpty(name))
                .Distinct()
                .OrderBy(name => name)
                .ToList();


            _selectionFunctions.UpdateCheckedListBox(modMain.arrListOfLOOM, checkedLoomItemsCL, chklistLoomCL);
            // _selectionFunctions.TotalCount(totalloom);
            _selectionFunctions.UpdateCheckedListBox(modMain.arrListOfSHEET, checkedSheetItemsCL, chkListSheetCL);

            _selectionFunctions.UpdateCheckedListBox(modMain.arrListOfLOOM, checkedLoomItemsCon, chkListLoomCont);
            _selectionFunctions.UpdateCheckedListBox(modMain.arrListOfLOOM, checkedLoomItemsMeg, chkListLoomMeg);
            _selectionFunctions.UpdateCheckedListBox(modMain.arrListOfSHEET, checkedSheetItemsPower, chkListSheetPowerOn);
        }
    }
}
