using Electre_Customize_DotNet.Contracts;
using Electre_Customize_DotNet.MainOperation;
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
    public partial class ContinuityWindow : UserControl
    {
        private HashSet<string> checkedLoomItems;
        private ISelectionFunctions _selectionFunctions;
        public int TotalLooms = modMain.arrListOfLOOM.Count();

        #region singleton
        private static ContinuityWindow? _continuityWindowInst;
        private static object lockObjet = new object();

        public static ContinuityWindow ContinuityWindowInst
        {
            get
            {
                if (_continuityWindowInst == null)
                {
                    lock (lockObjet)
                    {
                        if (_continuityWindowInst == null)
                        {
                            _continuityWindowInst = new ContinuityWindow();
                        }
                    }
                }
                return _continuityWindowInst;
            }
        }
        #endregion

        public List<string> SelectedLoomList
        {
            get
            {
                return checkedLoomItems.ToList();
            }
        }
        private ContinuityWindow()
        {
            InitializeComponent();
            checkedLoomItems = new HashSet<string>();
            _selectionFunctions = SelectionFunctions.SelectionFuntionsInstance;
           // modMain.arrListOfLOOM = modMain.arrListOfLOOM.OrderBy(x => x).ToList();
            _selectionFunctions.UpdateCheckedListBox(modMain.arrListOfLOOM, checkedLoomItems, chkListLoom);
            lblTotalLoom.Text = "Total Loom: " + TotalLooms.ToString();
        }

        // Commented on June 9th before implementing SelectAll intermediate functionality
        /* private void chkLoom_CheckedChanged(object sender, EventArgs e)
         {
             _selectionFunctions.SelectCheckItems(chkLoom, chkListLoom, checkedLoomItems);
         }*/
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

        // Commented on June 9th before implementing SelectAll intermediate functionality
        /* private void chkListLoom_ItemCheck(object sender, ItemCheckEventArgs e)
         {
             _selectionFunctions.ListItemChecked(e, checkedLoomItems, chkListLoom, lblLoomCount);
         }*/
        private void chkListLoom_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            _selectionFunctions.ListItemChecked(e, checkedLoomItems, chkListLoom, lblLoomCount, chkLoom);
        }
    }
}
