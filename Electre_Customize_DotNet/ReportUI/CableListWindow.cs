using Electre_Customize_DotNet.Contracts;
using Electre_Customize_DotNet.MainOperation;
using Electre_Customize_DotNet.Objects;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;
using Range = Microsoft.Office.Interop.Excel.Range;
using Electre_Customize_DotNet.Forms; // For drawing on the PDF

namespace Electre_Customize_DotNet.ReportUI
{
    public partial class CableListWindow : UserControl
    {
        private HashSet<string> checkedLoomItems;
        private HashSet<string> checkedSheetItems;
        private ISelectionFunctions _selectionFunctions;
        private static CableListWindow? _cableListWindowInst;
        public int countofloom = modMain.arrListOfLOOM.Count();
        public int countofsheet = modMain.arrListOfSHEET.Count();

        private static object lockObjet = new object();

        public static CableListWindow CableListWindowInst
        {
            get
            {
                if (_cableListWindowInst == null)
                {
                    lock (lockObjet)
                    {
                        if (_cableListWindowInst == null)
                        {
                            _cableListWindowInst = new CableListWindow();
                        }
                    }
                }

                return _cableListWindowInst;
            }
        }

        public bool ProjWirelist
        {
            get
            {
                if (chkProjSpec.Checked)
                    return true;

                else return false;
            }
        }

        public List<string> SelectedLoomList
        {
            get
            {
                return checkedLoomItems.ToList();
            }
        }

        public List<string> SelectedSheetList
        {
            get
            {
                return checkedSheetItems.ToList();
            }
        }

        private CableListWindow()
        {
            InitializeComponent();
            checkedLoomItems = new HashSet<string>();//keep the selected Items.
            checkedSheetItems = new HashSet<string>();
            _selectionFunctions = SelectionFunctions.SelectionFuntionsInstance;
           // modMain.arrListOfLOOM = modMain.arrListOfLOOM.OrderBy(x => x).ToList();
           // modMain.arrListOfSHEET = modMain.arrListOfSHEET.OrderBy(x => x).ToList();
            _selectionFunctions.UpdateCheckedListBox(modMain.arrListOfLOOM, checkedLoomItems, chkListLoom);
            // _selectionFunctions.TotalCount(totalloom);
            _selectionFunctions.UpdateCheckedListBox(modMain.arrListOfSHEET, checkedSheetItems, chkListSheet);
            // _selectionFunctions.TotalCount(totalsheet);

            lbltotalLoom.Text = "Total Looms: " + countofloom.ToString();
            lbltotalsheet.Text = "Total Sheets: " + countofsheet.ToString();
        }

        private void chkLoom_CheckedChanged(object sender, EventArgs e)
        {
            if (chkLoom.CheckState != CheckState.Indeterminate)
            {
                _selectionFunctions.SelectCheckItems(chkLoom, chkListLoom, checkedLoomItems);
            }
        }

        private void txtLoomFilter_TextChanged(object sender, EventArgs e)
        {
            _selectionFunctions.TextFilter(txtLoomFilter, checkedLoomItems, chkListLoom, modMain.arrListOfLOOM);
        }

        private void txtLoomFilter_Click(object sender, EventArgs e)
        {
            _selectionFunctions.TextClear(txtLoomFilter);
        }

        private void txtLoomFilter_Leave(object sender, EventArgs e)
        {
            _selectionFunctions.DefualtFilterName(txtLoomFilter);
        }

        private void chkListLoom_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            _selectionFunctions.ListItemChecked(e, checkedLoomItems, chkListLoom, lblLoomCount, chkLoom);
        }

        private void chkSheet_CheckedChanged(object sender, EventArgs e)
        {
            if (chkSheet.CheckState != CheckState.Indeterminate)
            {
                _selectionFunctions.SelectCheckItems(chkSheet, chkListSheet, checkedSheetItems);
            }
        }
        private void txtSheetFilter_TextChanged(object sender, EventArgs e)
        {
            _selectionFunctions.TextFilter(txtSheetFilter, checkedSheetItems, chkListSheet, modMain.arrListOfSHEET);
        }

        private void txtSheetFilter_Click(object sender, EventArgs e)
        {
            _selectionFunctions.TextClear(txtSheetFilter);
        }

        private void txtSheetFilter_Leave(object sender, EventArgs e)
        {
            _selectionFunctions.DefualtFilterName(txtSheetFilter);
        }

        private void chkListSheet_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            _selectionFunctions.ListItemChecked(e, checkedSheetItems, chkListSheet, lblSheetCount, chkSheet);
        }

        private void exportpdf_Click(object sender, EventArgs e)
        {
            // Create a new form to act as the pop-up window
            Form exportForm = new Form();

            //    // Create the ExportToPDF control
            ExportToPDF exportToPDF = new ExportToPDF();

            //    // Add the ExportToPDF control to the new form
            exportForm.Controls.Add(exportToPDF);

            //    // Set the ExportToPDF control to fill the entire form
            exportToPDF.Dock = DockStyle.Fill;

            //    // Set the pop-up form properties
            exportForm.Text = "Export PDF"; // Optional: Set the title of the pop-up
            exportForm.StartPosition = FormStartPosition.CenterScreen; // Optional: Center the pop-up on the screen
            exportForm.Size = new Size(1000, 400); // Optional: Set the size of the pop-up

            exportToPDF.PdfLocationTextBox.Text = GlobalVar.ReportFolderGlobal + ConfigurationManager.AppSettings["BundleFolder"];
            //    // Show the form as a separate window
            exportForm.ShowDialog();

        }
    }
}
