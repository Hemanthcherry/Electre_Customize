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
    public partial class ComponentBreak : UserControl
    {

        private HashSet<string> checkedEqupItems;
        private HashSet<string> checkedJmItems;
        private HashSet<string> checkedBrkItems;
        private HashSet<string> checkedMiscItems;
        public int TotalEqu = modMain.arrListOfEQU.Count();
        public int TotalBreakconnectors = modMain.arrListOfDIS.Count();
        //public int Totaljunctionandsplice = modMain.arrListOfJUN.Count() + modMain.arrListOfJUNnSPL.Count();
        public int Totaljunctionandsplice = modMain.arrListOfJUNnSPL.Count();
        public int Totalmisc = modMain.arrListOfSWT.Count() + modMain.arrListOfREL.Count();

        //singletone initialization
        #region singletone
        private static ComponentBreak? _componentBreakInst;
        private static object lockObjet = new object();

        public static ComponentBreak ComponentBreakInst
        {
            get
            {
                if (_componentBreakInst == null)
                {
                    lock (lockObjet)
                    {
                        if (_componentBreakInst == null)
                        {
                            _componentBreakInst = new ComponentBreak();
                        }
                    }
                }
                return _componentBreakInst;
            }
        }
        #endregion

        private ISelectionFunctions _selectionFunctions;

        public List<string> SelectedEqupList
        {
            get
            {
                return checkedEqupItems.ToList();
            }
        }

        public List<string> SelectedBrkList
        {
            get
            {
                return checkedBrkItems.ToList();
            }
        }

        public List<string> SelectedJmList
        {
            get
            {
                return checkedJmItems.ToList();
            }
        }

        public List<string> SelectedMiscList
        {
            get
            {
                return checkedMiscItems.ToList();
            }
        }

        private ComponentBreak()
        {
            InitializeComponent();
            checkedEqupItems = new HashSet<string>();
            checkedMiscItems = new HashSet<string>();
            checkedBrkItems = new HashSet<string>();
            checkedJmItems = new HashSet<string>();
            modMain.arrListOfJUNnSPL = modMain.arrListOfJUNnSPL.Where(item => !string.IsNullOrWhiteSpace(item)).ToList();
            _selectionFunctions = SelectionFunctions.SelectionFuntionsInstance;
            _selectionFunctions.UpdateCheckedListBox(modMain.arrListOfEQU, checkedEqupItems, chkListEqup);
            _selectionFunctions.UpdateCheckedListBox(modMain.arrListOfDIS, checkedBrkItems, chkListBrk);
            //_selectionFunctions.UpdateCheckedListBox(modMain.arrListOfJUN.Concat(modMain.arrListOfJUNnSPL).ToList(), checkedJmItems, chkListJm);
            _selectionFunctions.UpdateCheckedListBox(modMain.arrListOfJUNnSPL, checkedJmItems, chkListJm);
            _selectionFunctions.UpdateCheckedListBox(modMain.arrListOfREL.Concat(modMain.arrListOfSWT).ToList(), checkedMiscItems, chkListMisc);
            lblTotalEque.Text = "Total Equipments: " + TotalEqu.ToString();
            lblTotalBreckcon.Text = "Total Break Connectors: " + TotalBreakconnectors.ToString();
            lblTotaljun.Text = "Total junction Modules: " + Totaljunctionandsplice.ToString();
            lblTotalmis.Text = "Total misc: " + Totalmisc.ToString();
        }

        #region Equipment

        private void chkEqup_CheckedChanged(object sender, EventArgs e)
        {
            if (chkEqup.CheckState != CheckState.Indeterminate)
            {
                _selectionFunctions.SelectCheckItems(chkEqup, chkListEqup, checkedEqupItems);
            }
        }

        private void txtEqupFilter_TextChanged(object sender, EventArgs e)
        {
            _selectionFunctions.TextFilter(txtEqupFilter, checkedEqupItems, chkListEqup, modMain.arrListOfEQU);
        }

        private void txtEqupFilter_Click(object sender, EventArgs e)
        {
            _selectionFunctions.TextClear(txtEqupFilter);
        }

        private void txtEqupFilter_Leave(object sender, EventArgs e)
        {
            _selectionFunctions.DefualtFilterName(txtEqupFilter);
        }

        private void chkListEqup_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            _selectionFunctions.ListItemChecked(e, checkedEqupItems, chkListEqup, lblEqupCount,chkEqup);
        }

        #endregion
        #region BreakConnector
        private void chkBrk_CheckedChanged(object sender, EventArgs e)
        {
            if (chkBrk.CheckState != CheckState.Indeterminate)
            {
                _selectionFunctions.SelectCheckItems(chkBrk, chkListBrk, checkedBrkItems);
            }
        }

        private void txtBrkFilter_TextChanged(object sender, EventArgs e)
        {
            _selectionFunctions.TextFilter(txtBrkFilter, checkedBrkItems, chkListBrk, modMain.arrListOfDIS);

        }

        private void txtBrkFilter_Click(object sender, EventArgs e)
        {
            _selectionFunctions.TextClear(txtBrkFilter);
        }

        private void txtBrkFilter_Leave(object sender, EventArgs e)
        {
            _selectionFunctions.DefualtFilterName(txtBrkFilter);
        }

        private void chkListBrk_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            _selectionFunctions.ListItemChecked(e, checkedBrkItems, chkListBrk, lblBrkCount, chkBrk);
        }

        #endregion
        #region Junction Module

        private void chkJm_CheckedChanged(object sender, EventArgs e)
        {
            if (chkJm.CheckState != CheckState.Indeterminate)
            {
                _selectionFunctions.SelectCheckItems(chkJm, chkListJm, checkedJmItems);
            }
        }

        private void txtJmFilter_TextChanged(object sender, EventArgs e)
        {
            _selectionFunctions.TextFilter(txtJmFilter, checkedJmItems, chkListJm, modMain.arrListOfJUNnSPL.ToList());
        }

        private void txtJmFilter_Click(object sender, EventArgs e)
        {
            _selectionFunctions.TextClear(txtJmFilter);
        }

        private void txtJmFilter_Leave(object sender, EventArgs e)
        {
            _selectionFunctions.DefualtFilterName(txtJmFilter);
        }

        private void chkListJm_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            _selectionFunctions.ListItemChecked(e, checkedJmItems, chkListJm, lblJmCount,chkJm);
        }
        #endregion

        #region MISC

        private void chkMisc_CheckedChanged(object sender, EventArgs e)
        {
            if (chkMisc.CheckState != CheckState.Indeterminate)
            {
                _selectionFunctions.SelectCheckItems(chkMisc, chkListMisc, checkedMiscItems);
            }
        }

        private void txtMiscFilter_TextChanged(object sender, EventArgs e)
        {
            _selectionFunctions.TextFilter(txtMiscFilter, checkedMiscItems, chkListMisc, modMain.arrListOfREL.Concat(modMain.arrListOfSWT).ToList());
        }

        private void txtMiscFilter_Click(object sender, EventArgs e)
        {
            _selectionFunctions.TextClear(txtMiscFilter);
        }

        private void txtMiscFilter_Leave(object sender, EventArgs e)
        {
            _selectionFunctions.DefualtFilterName(txtMiscFilter);
        }

        private void chkListMisc_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            _selectionFunctions.ListItemChecked(e, checkedMiscItems, chkListMisc, lblMiscCount,chkMisc);
        }
        #endregion

    }
}
