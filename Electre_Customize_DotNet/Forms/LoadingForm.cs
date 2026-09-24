using System.Diagnostics;
using System.Reflection;

namespace Electre_Customize_DotNet
{
    public partial class LoadingForm : Form
    {

        //old commented on Jan,06,2025

        //public string LbsLoadinMessag
        //{
        //    get
        //    { return lableLoadMessage.Text; }

        //    set
        //    {
        //        if (string.IsNullOrEmpty(lableLoadMessage.Text))
        //        {
        //            lableLoadMessage.Text = value;
        //        }

        //        else
        //        {
        //            lableLoadMessage.Text = $"{lableLoadMessage.Text}\n" + value;
        //        }
        //    }
        //}

        public string LbsLoadinMessag
        {
            get
            {
                if (lableLoadMessage.InvokeRequired)
                    return (string)lableLoadMessage.Invoke(
                        new Func<string>(() => lableLoadMessage.Text));

                return lableLoadMessage.Text;
            }
            set
            {
                if (lableLoadMessage.InvokeRequired)
                {
                    lableLoadMessage.Invoke(new Action(() => LbsLoadinMessag = value));
                    return;
                }

                if (string.IsNullOrEmpty(lableLoadMessage.Text))
                {
                    lableLoadMessage.Text = value;
                }
                else
                {
                    lableLoadMessage.Text += Environment.NewLine + value;
                }
            }
        }

        //old commented on Jan,06,2025
        //public string CompletionMessage
        //{
        //    get
        //    { return lableLoadMessage.Text; }

        //    set
        //    {
        //        if (string.IsNullOrEmpty(completionlbl.Text))
        //        {
        //            completionlbl.Text = value;
        //        }

        //        else
        //        {
        //            completionlbl.Text = value;
        //        }
        //    }
        //}

        public string CompletionMessage
        {
            get
            {
                if (completionlbl.InvokeRequired)
                    return (string)completionlbl.Invoke(
                        new Func<string>(() => completionlbl.Text));

                return completionlbl.Text;
            }
            set
            {
                if (completionlbl.InvokeRequired)
                {
                    completionlbl.Invoke(new Action(() => CompletionMessage = value));
                    return;
                }

                completionlbl.Text = value;
            }
        }


        public string Loadingbtn
        {
            get
            {
                if (loadingbtn.InvokeRequired)
                    return (string)loadingbtn.Invoke(new Func<string>(() => loadingbtn.Text));

                return loadingbtn.Text;
            }
            set
            {
                if (loadingbtn.InvokeRequired)
                {
                    loadingbtn.Invoke(new Action(() => Loadingbtn = value));
                    return;
                }

                loadingbtn.Text = value;

                if (value.Equals("wait", StringComparison.OrdinalIgnoreCase))
                {
                    loadingbtn.BackColor = Color.Yellow;
                }
                else
                {
                    loadingbtn.BackColor = Color.GreenYellow;
                }
            }
        }

        //old commented on Jan,06,2025
        //public string Loadingbtn
        //{
        //    get { return loadingbtn.Text; }

        //    set
        //    {
        //        if (string.Equals(value, "wait", StringComparison.OrdinalIgnoreCase))
        //        {
        //            loadingbtn.Text = value;
        //            loadingbtn.BackColor = Color.Yellow;
        //        }

        //        else
        //        {
        //            loadingbtn.Text = value;
        //            loadingbtn.BackColor = Color.GreenYellow;
        //        }
        //    }
        //}

        public bool LoadingbtnEnable
        {
            get
            {
                if (loadingbtn.InvokeRequired)
                    return (bool)loadingbtn.Invoke(new Func<bool>(() => loadingbtn.Enabled));

                return loadingbtn.Enabled;
            }
            set
            {
                if (loadingbtn.InvokeRequired)
                {
                    loadingbtn.Invoke(new Action(() => LoadingbtnEnable = value));
                    return;
                }

                loadingbtn.Enabled = value;
            }
        }

        //old commented on Jan,06,2025
        //public bool LoadingbtnEnable
        //{
        //    get { return loadingbtn.Enabled; }
        //    set { loadingbtn.Enabled = value; }
        //}

        public string ReportLoclbl
        {
            get
            {
                if (reportLoclbl.InvokeRequired)
                    return (string)reportLoclbl.Invoke(
                        new Func<string>(() => reportLoclbl.Text));

                return reportLoclbl.Text;
            }
            set
            {
                if (reportLoclbl.InvokeRequired)
                {
                    reportLoclbl.Invoke(new Action(() => ReportLoclbl = value));
                    return;
                }

                reportLocLinklbl.Show();
                reportLoclbl.Text = value;
            }
        }

        //old commented on Jan,06,2025
        //public string ReportLoclbl
        //{
        //    get { return reportLoclbl.ToString(); }
        //    set
        //    {
        //        reportLocLinklbl.Show();
        //        reportLoclbl.Text = value;
        //    }
        //}

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

            if (string.Equals(loadingbtn.Text, "close", StringComparison.OrdinalIgnoreCase))
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
