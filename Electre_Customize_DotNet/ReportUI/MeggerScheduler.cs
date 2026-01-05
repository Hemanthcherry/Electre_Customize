using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Electre_Customize_DotNet.Contracts;
using Electre_Customize_DotNet.MainOperation;
using System.Threading.Tasks;
using System.Diagnostics.Eventing.Reader;

namespace Electre_Customize_DotNet.ReportUI
{
    public partial class MeggerScheduler : UserControl
    {
        private HashSet<string> checkedLoomItems;
        private ISelectionFunctions _selectionFunctions;
        public int TotalLooms = modMain.arrListOfLOOM.Count();

        public static MeggerScheduler _meggerschedulerinst;
        private static object lockobject = new object();
        public static MeggerScheduler MeggerSchedulerinst
        {
            get
            {
                if (_meggerschedulerinst == null)
                {
                    lock (lockobject)
                    {
                        if (_meggerschedulerinst == null)
                        {
                            _meggerschedulerinst = new MeggerScheduler();
                        }
                    }
                }

                return _meggerschedulerinst;
            }
        }
/*        public bool Meggerscheduler
        {
            get
            {
                if (Megger.Checked)

                    return true;


                else return false;

            }
        }
*/        public List<string> SelectedLoomList
        {
            get
            {
                return checkedLoomItems.ToList();
            }
        }
        public MeggerScheduler()
        {
            InitializeComponent();
            checkedLoomItems = new HashSet<string>();
            _selectionFunctions = SelectionFunctions.SelectionFuntionsInstance;
           // modMain.arrListOfLOOM = modMain.arrListOfLOOM.OrderBy(x => x).ToList();
            _selectionFunctions.UpdateCheckedListBox(modMain.arrListOfLOOM, checkedLoomItems, chkListLoom);
            labelTotal.Text = "Total Looms: " + TotalLooms.ToString();
        }

      /*  private void MeggerScheduler_Load(object sender, EventArgs e)
        {

        }*/

       /* private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }*/

        private void textLoomFilter_TextChanged(object sender, EventArgs e)
        {
            _selectionFunctions.TextFilter(txtLoomFilter, checkedLoomItems, chkListLoom, modMain.arrListOfLOOM);
        }

        private void checkLoom_CheckedChanged(object sender, EventArgs e)
        {
            if (chkLoom.CheckState != CheckState.Indeterminate)
            {
                _selectionFunctions.SelectCheckItems(chkLoom, chkListLoom, checkedLoomItems);
            }
        }

        private void checkedListLoom_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
        private void checkedListLoom_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            _selectionFunctions.ListItemChecked(e, checkedLoomItems, chkListLoom, lblLoomCount,chkLoom);
        }
        private void textLoomFilter_Click(object sender, EventArgs e)
        {
            _selectionFunctions.TextClear(txtLoomFilter);
        }
        private void textLoomFilter_Leave(object sender, EventArgs e)
        {
            _selectionFunctions.DefualtFilterName(txtLoomFilter);
        }
    }
}
