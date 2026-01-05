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
    public partial class CustomMessageBox : Form
    {
        public enum CustomResult
        {
            WithConfig,
            WithoutConfig
        }

        public CustomResult Result { get; private set; }

        public CustomMessageBox(string message, string title)
        {
            InitializeComponent();
            this.Text = title;
            lblMessage.Text = message;
        }

        private void btnWithConfig_Click(object sender, EventArgs e)
        {
            Result = CustomResult.WithConfig;
            this.DialogResult = DialogResult.OK;
            Close();
        }

        private void btnWithoutConfig_Click(object sender, EventArgs e)
        {
            Result = CustomResult.WithoutConfig;
            this.DialogResult = DialogResult.OK;
            Close();
        }
    }
}
