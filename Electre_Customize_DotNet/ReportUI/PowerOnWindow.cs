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
    public partial class PowerOnWindow : UserControl
    {
        // private HashSet<string> checkedLoomItems;
        private HashSet<string> _checkedPanels;
        private ISelectionFunctions _selectionFunctions;

        #region singletone
        private static PowerOnWindow? _powerOnWindowInst;
        private static object lockObjet = new object();

        public static PowerOnWindow PowerOnWindowInst
        {
            get
            {
                if (_powerOnWindowInst == null)
                {
                    lock (lockObjet)
                    {
                        if (_powerOnWindowInst == null)
                        {
                            _powerOnWindowInst = new PowerOnWindow();
                        }
                    }
                }
                return _powerOnWindowInst;
            }
        }
        #endregion

        public List<string> SelectedPanelList
        {
            get
            {
                return _checkedPanels.ToList();
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
        private PowerOnWindow()
        {
            InitializeComponent();
            _checkedPanels = new HashSet<string>();
            _selectionFunctions = SelectionFunctions.SelectionFuntionsInstance;
            modMain.arrListOfPANEL = modMain.arrListOfPANEL.Where(item => !string.IsNullOrWhiteSpace(item)).ToList();
            _selectionFunctions.UpdateCheckedListBox(modMain.arrListOfPANEL, _checkedPanels, chkListPanel);
            lblTotalPanel.Text = "Total Panels: " + modMain.arrListOfPANEL.Count.ToString();
        }

         private void chkPanel_CheckedChanged(object sender, EventArgs e)
         {
              if (chkPanel.CheckState != CheckState.Indeterminate)
              {
                  _selectionFunctions.SelectCheckItems(chkPanel, chkListPanel, _checkedPanels);
              }
          }

        private void txtPanelFilter_TextChanged(object sender, EventArgs e)
        {
            _selectionFunctions.TextFilter(txtPanelFilter, _checkedPanels, chkListPanel, modMain.arrListOfPANEL);
        }

        private void txtPanelFilter_Click(object sender, EventArgs e)
        {
            _selectionFunctions.TextClear(txtPanelFilter);
        }

        private void txtPanelFilter_Leave(object sender, EventArgs e)
        {
            _selectionFunctions.DefualtFilterName(txtPanelFilter);
        }

        private void chkListPanel_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            _selectionFunctions.ListItemChecked(e, _checkedPanels, chkListPanel, lblPanelCount, chkPanel);
        }

        /*private void chkListPanel_SelectedIndexChanged(object sender, EventArgs e)
        {

        }*/
    }
}

