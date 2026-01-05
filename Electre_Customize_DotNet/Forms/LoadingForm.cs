using System.Diagnostics;
using System.Reflection;

namespace Electre_Customize_DotNet
{
    public partial class LoadingForm : Form
    {
        public string LbsLoadinMessag
        {
            get
            { return lableLoadMessage.Text; }

            set
            {
                if (string.IsNullOrEmpty(lableLoadMessage.Text))
                {
                    lableLoadMessage.Text = value;
                }

                else
                {
                    lableLoadMessage.Text = $"{lableLoadMessage.Text}\n" + value;
                }
            }
        }

        public string CompletionMessage
        {
            get
            { return lableLoadMessage.Text; }

            set
            {
                if (string.IsNullOrEmpty(completionlbl.Text))
                {
                    completionlbl.Text = value;
                }

                else
                {
                    completionlbl.Text = value;
                }
            }
        }

        public string Loadingbtn
        {
            get { return loadingbtn.Text; }

            set
            {

                if (value.ToLower() == "wait")
                {
                    loadingbtn.Text = value;
                    loadingbtn.BackColor = Color.Yellow;
                }

                else
                {
                    loadingbtn.Text = value;
                    loadingbtn.BackColor = Color.GreenYellow;
                }
            }
        }

        public bool LoadingbtnEnable
        {
            get { return loadingbtn.Enabled; }
            set { loadingbtn.Enabled = value; }
        }

        public string ReportLoclbl
        {
            get { return reportLoclbl.ToString(); }
            set
            {
                reportLocLinklbl.Show();
                reportLoclbl.Text = value;
            }
        }




        public LoadingForm()
        {
            InitializeComponent();
        }

        private void Loadingbtn_Click(object sender, EventArgs e)
        {
            lableLoadMessage.Text = "";
            this.Close();
        }

      





        private void LoadingForm_Load(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            reportLocLinklbl.Hide();

            if (loadingbtn.Text.ToLower() == "close")
            {
                Cursor = Cursors.Default;
            }
        }

        private void reportLocLinklbl_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            string[] reportLocation = reportLoclbl.Text.Split('@');

            if (!string.IsNullOrEmpty(reportLocation[1]))
            {
                reportLocation[1].Trim();

                if (Directory.Exists(reportLocation[1]))
                {
                    Process.Start("explorer.exe", reportLocation[1]);
                }
            }
        }






    }
}
