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

namespace Electre_Customize_DotNet.Forms
{
    public partial class MessageBoxWindow : Form
    {

        List<DuplicateWire> duplicateWires = new List<DuplicateWire>();
        public MessageBoxWindow(List<DuplicateWire> duplicateWire)
        {
            InitializeComponent();
            this.duplicateWires = duplicateWire;
        }

        private void MessageBoxWindow_Load(object sender, EventArgs e)
        {
            duplicateWireGridview.DataSource = duplicateWires;
        }

        private void wireListBtn_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void wireList_Click(object sender, EventArgs e)
        {

        }

        private void duplicateWireGridview_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
