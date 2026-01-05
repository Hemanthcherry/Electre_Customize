namespace Electre_Customize_DotNet.ReportUI
{
    partial class ComponentBreak
    {
        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            chkEqup = new CheckBox();
            txtEqupFilter = new TextBox();
            chkListEqup = new CheckedListBox();
            groupBox1 = new GroupBox();
            lblTotalEque = new Label();
            lblEqupCount = new Label();
            groupBox2 = new GroupBox();
            lblTotalBreckcon = new Label();
            lblBrkCount = new Label();
            txtBrkFilter = new TextBox();
            chkBrk = new CheckBox();
            chkListBrk = new CheckedListBox();
            groupBox3 = new GroupBox();
            lblTotalmis = new Label();
            lblMiscCount = new Label();
            txtMiscFilter = new TextBox();
            chkMisc = new CheckBox();
            chkListMisc = new CheckedListBox();
            groupBox4 = new GroupBox();
            lblTotaljun = new Label();
            lblJmCount = new Label();
            txtJmFilter = new TextBox();
            chkJm = new CheckBox();
            chkListJm = new CheckedListBox();
            groupBox1.SuspendLayout();
            groupBox2.SuspendLayout();
            groupBox3.SuspendLayout();
            groupBox4.SuspendLayout();
            SuspendLayout();
            // 
            // chkEqup
            // 
            chkEqup.AutoSize = true;
            chkEqup.Cursor = Cursors.Hand;
            chkEqup.Font = new Font("Segoe UI", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0);
            chkEqup.Location = new Point(300, 62);
            chkEqup.Margin = new Padding(4, 5, 4, 5);
            chkEqup.Name = "chkEqup";
            chkEqup.Size = new Size(127, 32);
            chkEqup.TabIndex = 37;
            chkEqup.Text = "Select All";
            chkEqup.UseVisualStyleBackColor = true;
            chkEqup.CheckedChanged += chkEqup_CheckedChanged;
            // 
            // txtEqupFilter
            // 
            txtEqupFilter.Font = new Font("Segoe UI", 12F, FontStyle.Italic);
            txtEqupFilter.ForeColor = Color.Silver;
            txtEqupFilter.Location = new Point(9, 55);
            txtEqupFilter.Margin = new Padding(4, 5, 4, 5);
            txtEqupFilter.Name = "txtEqupFilter";
            txtEqupFilter.Size = new Size(273, 39);
            txtEqupFilter.TabIndex = 39;
            txtEqupFilter.Text = "Search Filter";
            txtEqupFilter.Click += txtEqupFilter_Click;
            txtEqupFilter.TextChanged += txtEqupFilter_TextChanged;
            txtEqupFilter.Leave += txtEqupFilter_Leave;
            // 
            // chkListEqup
            // 
            chkListEqup.CheckOnClick = true;
            chkListEqup.Font = new Font("Consolas", 11F, FontStyle.Regular, GraphicsUnit.Point, 0);
            chkListEqup.FormattingEnabled = true;
            chkListEqup.Location = new Point(9, 113);
            chkListEqup.Margin = new Padding(4, 5, 4, 5);
            chkListEqup.Name = "chkListEqup";
            chkListEqup.ScrollAlwaysVisible = true;
            chkListEqup.Size = new Size(410, 214);
            chkListEqup.TabIndex = 38;
            chkListEqup.ItemCheck += chkListEqup_ItemCheck;
            //chkListEqup.SelectedIndexChanged += chkListEqup_SelectedIndexChanged;
            // 
            // groupBox1
            // 
            groupBox1.Controls.Add(lblTotalEque);
            groupBox1.Controls.Add(lblEqupCount);
            groupBox1.Controls.Add(txtEqupFilter);
            groupBox1.Controls.Add(chkEqup);
            groupBox1.Controls.Add(chkListEqup);
            groupBox1.Font = new Font("Segoe UI", 12F, FontStyle.Bold, GraphicsUnit.Point, 0);
            groupBox1.Location = new Point(7, 8);
            groupBox1.Margin = new Padding(7, 8, 7, 8);
            groupBox1.Name = "groupBox1";
            groupBox1.Padding = new Padding(4, 5, 4, 5);
            groupBox1.Size = new Size(429, 400);
            groupBox1.TabIndex = 40;
            groupBox1.TabStop = false;
            groupBox1.Text = "Equipment";
            // 
            // lblTotalEque
            // 
            lblTotalEque.AutoSize = true;
            lblTotalEque.Font = new Font("Segoe UI Semibold", 9F, FontStyle.Bold | FontStyle.Italic, GraphicsUnit.Point, 0);
            lblTotalEque.ForeColor = Color.Blue;
            lblTotalEque.Location = new Point(223, 357);
            lblTotalEque.Name = "lblTotalEque";
            lblTotalEque.Size = new Size(170, 25);
            lblTotalEque.TabIndex = 42;
            lblTotalEque.Text = "Total Equipments: 0";
           // lblTotalEque.Click += lblTotalEque_Click;
            // 
            // lblEqupCount
            // 
            lblEqupCount.AutoSize = true;
            lblEqupCount.Font = new Font("Segoe UI Semibold", 9F, FontStyle.Bold | FontStyle.Italic, GraphicsUnit.Point, 0);
            lblEqupCount.ForeColor = Color.Blue;
            lblEqupCount.Location = new Point(197, 332);
            lblEqupCount.Margin = new Padding(4, 0, 4, 0);
            lblEqupCount.Name = "lblEqupCount";
            lblEqupCount.Size = new Size(197, 25);
            lblEqupCount.TabIndex = 41;
            lblEqupCount.Text = "Selected Items Count: 0";
           // lblEqupCount.Click += lblEqupCount_Click;
            // 
            // groupBox2
            // 
            groupBox2.Controls.Add(lblTotalBreckcon);
            groupBox2.Controls.Add(lblBrkCount);
            groupBox2.Controls.Add(txtBrkFilter);
            groupBox2.Controls.Add(chkBrk);
            groupBox2.Controls.Add(chkListBrk);
            groupBox2.Font = new Font("Segoe UI", 12F, FontStyle.Bold, GraphicsUnit.Point, 0);
            groupBox2.Location = new Point(457, 8);
            groupBox2.Margin = new Padding(7, 8, 7, 8);
            groupBox2.Name = "groupBox2";
            groupBox2.Padding = new Padding(4, 5, 4, 5);
            groupBox2.Size = new Size(429, 400);
            groupBox2.TabIndex = 41;
            groupBox2.TabStop = false;
            groupBox2.Text = "Break Connector";
            // 
            // lblTotalBreckcon
            // 
            lblTotalBreckcon.AutoSize = true;
            lblTotalBreckcon.Font = new Font("Segoe UI Semibold", 9F, FontStyle.Bold | FontStyle.Italic, GraphicsUnit.Point, 0);
            lblTotalBreckcon.ForeColor = Color.Blue;
            lblTotalBreckcon.Location = new Point(174, 357);
            lblTotalBreckcon.Name = "lblTotalBreckcon";
            lblTotalBreckcon.Size = new Size(209, 25);
            lblTotalBreckcon.TabIndex = 43;
            lblTotalBreckcon.Text = "Total Break Connector: 0";
            // 
            // lblBrkCount
            // 
            lblBrkCount.AutoSize = true;
            lblBrkCount.Font = new Font("Segoe UI Semibold", 9F, FontStyle.Bold | FontStyle.Italic, GraphicsUnit.Point, 0);
            lblBrkCount.ForeColor = Color.Blue;
            lblBrkCount.Location = new Point(186, 332);
            lblBrkCount.Margin = new Padding(4, 0, 4, 0);
            lblBrkCount.Name = "lblBrkCount";
            lblBrkCount.Size = new Size(197, 25);
            lblBrkCount.TabIndex = 42;
            lblBrkCount.Text = "Selected Items Count: 0";
           // lblBrkCount.Click += lblBrkCount_Click;
            // 
            // txtBrkFilter
            // 
            txtBrkFilter.Font = new Font("Segoe UI", 12F, FontStyle.Italic);
            txtBrkFilter.ForeColor = Color.Silver;
            txtBrkFilter.Location = new Point(9, 55);
            txtBrkFilter.Margin = new Padding(4, 5, 4, 5);
            txtBrkFilter.Name = "txtBrkFilter";
            txtBrkFilter.Size = new Size(273, 39);
            txtBrkFilter.TabIndex = 39;
            txtBrkFilter.Text = "Search Filter";
            txtBrkFilter.Click += txtBrkFilter_Click;
            txtBrkFilter.TextChanged += txtBrkFilter_TextChanged;
            txtBrkFilter.Leave += txtBrkFilter_Leave;
            // 
            // chkBrk
            // 
            chkBrk.AutoSize = true;
            chkBrk.Cursor = Cursors.Hand;
            chkBrk.Font = new Font("Segoe UI", 9.75F, FontStyle.Bold);
            chkBrk.Location = new Point(300, 62);
            chkBrk.Margin = new Padding(4, 5, 4, 5);
            chkBrk.Name = "chkBrk";
            chkBrk.Size = new Size(127, 32);
            chkBrk.TabIndex = 37;
            chkBrk.Text = "Select All";
            chkBrk.UseVisualStyleBackColor = true;
            chkBrk.CheckedChanged += chkBrk_CheckedChanged;
            // 
            // chkListBrk
            // 
            chkListBrk.CheckOnClick = true;
            chkListBrk.Font = new Font("Consolas", 11F);
            chkListBrk.FormattingEnabled = true;
            chkListBrk.Location = new Point(9, 113);
            chkListBrk.Margin = new Padding(4, 5, 4, 5);
            chkListBrk.Name = "chkListBrk";
            chkListBrk.ScrollAlwaysVisible = true;
            chkListBrk.Size = new Size(410, 214);
            chkListBrk.TabIndex = 38;
            chkListBrk.ItemCheck += chkListBrk_ItemCheck;
            //chkListBrk.SelectedIndexChanged += chkListBrk_SelectedIndexChanged;
            // 
            // groupBox3
            // 
            groupBox3.Controls.Add(lblTotalmis);
            groupBox3.Controls.Add(lblMiscCount);
            groupBox3.Controls.Add(txtMiscFilter);
            groupBox3.Controls.Add(chkMisc);
            groupBox3.Controls.Add(chkListMisc);
            groupBox3.Font = new Font("Segoe UI", 12F, FontStyle.Bold, GraphicsUnit.Point, 0);
            groupBox3.Location = new Point(457, 425);
            groupBox3.Margin = new Padding(7, 8, 7, 8);
            groupBox3.Name = "groupBox3";
            groupBox3.Padding = new Padding(4, 5, 4, 5);
            groupBox3.Size = new Size(429, 400);
            groupBox3.TabIndex = 43;
            groupBox3.TabStop = false;
            groupBox3.Text = "MISC";
            // 
            // lblTotalmis
            // 
            lblTotalmis.AutoSize = true;
            lblTotalmis.Font = new Font("Segoe UI Semibold", 9F, FontStyle.Bold | FontStyle.Italic, GraphicsUnit.Point, 0);
            lblTotalmis.ForeColor = Color.Blue;
            lblTotalmis.Location = new Point(257, 358);
            lblTotalmis.Name = "lblTotalmis";
            lblTotalmis.Size = new Size(106, 25);
            lblTotalmis.TabIndex = 44;
            lblTotalmis.Text = "Total mis: 0";
            // 
            // lblMiscCount
            // 
            lblMiscCount.AutoSize = true;
            lblMiscCount.Font = new Font("Segoe UI Semibold", 9F, FontStyle.Bold | FontStyle.Italic, GraphicsUnit.Point, 0);
            lblMiscCount.ForeColor = Color.Blue;
            lblMiscCount.Location = new Point(166, 332);
            lblMiscCount.Margin = new Padding(4, 0, 4, 0);
            lblMiscCount.Name = "lblMiscCount";
            lblMiscCount.Size = new Size(197, 25);
            lblMiscCount.TabIndex = 43;
            lblMiscCount.Text = "Selected Items Count: 0";
           // lblMiscCount.Click += lblMiscCount_Click;
            // 
            // txtMiscFilter
            // 
            txtMiscFilter.Font = new Font("Segoe UI", 12F, FontStyle.Italic);
            txtMiscFilter.ForeColor = Color.Silver;
            txtMiscFilter.Location = new Point(9, 55);
            txtMiscFilter.Margin = new Padding(4, 5, 4, 5);
            txtMiscFilter.Name = "txtMiscFilter";
            txtMiscFilter.Size = new Size(273, 39);
            txtMiscFilter.TabIndex = 39;
            txtMiscFilter.Text = "Search Filter";
            txtMiscFilter.Click += txtMiscFilter_Click;
            txtMiscFilter.TextChanged += txtMiscFilter_TextChanged;
            txtMiscFilter.Leave += txtMiscFilter_Leave;
            // 
            // chkMisc
            // 
            chkMisc.AutoSize = true;
            chkMisc.Cursor = Cursors.Hand;
            chkMisc.Font = new Font("Segoe UI", 9.75F, FontStyle.Bold);
            chkMisc.Location = new Point(300, 62);
            chkMisc.Margin = new Padding(4, 5, 4, 5);
            chkMisc.Name = "chkMisc";
            chkMisc.Size = new Size(127, 32);
            chkMisc.TabIndex = 37;
            chkMisc.Text = "Select All";
            chkMisc.UseVisualStyleBackColor = true;
            chkMisc.CheckedChanged += chkMisc_CheckedChanged;
            // 
            // chkListMisc
            // 
            chkListMisc.CheckOnClick = true;
            chkListMisc.Font = new Font("Consolas", 11F);
            chkListMisc.FormattingEnabled = true;
            chkListMisc.Location = new Point(9, 113);
            chkListMisc.Margin = new Padding(4, 5, 4, 5);
            chkListMisc.Name = "chkListMisc";
            chkListMisc.ScrollAlwaysVisible = true;
            chkListMisc.Size = new Size(410, 214);
            chkListMisc.TabIndex = 38;
            chkListMisc.ItemCheck += chkListMisc_ItemCheck;
            // 
            // groupBox4
            // 
            groupBox4.Controls.Add(lblTotaljun);
            groupBox4.Controls.Add(lblJmCount);
            groupBox4.Controls.Add(txtJmFilter);
            groupBox4.Controls.Add(chkJm);
            groupBox4.Controls.Add(chkListJm);
            groupBox4.Font = new Font("Segoe UI", 12F, FontStyle.Bold, GraphicsUnit.Point, 0);
            groupBox4.Location = new Point(7, 425);
            groupBox4.Margin = new Padding(7, 8, 7, 8);
            groupBox4.Name = "groupBox4";
            groupBox4.Padding = new Padding(4, 5, 4, 5);
            groupBox4.Size = new Size(429, 400);
            groupBox4.TabIndex = 42;
            groupBox4.TabStop = false;
            groupBox4.Text = "Junction Module";
            // 
            // lblTotaljun
            // 
            lblTotaljun.AutoSize = true;
            lblTotaljun.Font = new Font("Segoe UI Semibold", 9F, FontStyle.Bold | FontStyle.Italic, GraphicsUnit.Point, 0);
            lblTotaljun.ForeColor = Color.Blue;
            lblTotaljun.Location = new Point(178, 358);
            lblTotaljun.Name = "lblTotaljun";
            lblTotaljun.Size = new Size(215, 25);
            lblTotaljun.TabIndex = 45;
            lblTotaljun.Text = "Total Junction Modules: 0";
            // 
            // lblJmCount
            // 
            lblJmCount.AutoSize = true;
            lblJmCount.Font = new Font("Segoe UI Semibold", 9F, FontStyle.Bold | FontStyle.Italic, GraphicsUnit.Point, 0);
            lblJmCount.ForeColor = Color.Blue;
            lblJmCount.Location = new Point(197, 332);
            lblJmCount.Margin = new Padding(4, 0, 4, 0);
            lblJmCount.Name = "lblJmCount";
            lblJmCount.Size = new Size(197, 25);
            lblJmCount.TabIndex = 44;
            lblJmCount.Text = "Selected Items Count: 0";
            // 
            // txtJmFilter
            // 
            txtJmFilter.Font = new Font("Segoe UI", 12F, FontStyle.Italic);
            txtJmFilter.ForeColor = Color.Silver;
            txtJmFilter.Location = new Point(9, 55);
            txtJmFilter.Margin = new Padding(4, 5, 4, 5);
            txtJmFilter.Name = "txtJmFilter";
            txtJmFilter.Size = new Size(273, 39);
            txtJmFilter.TabIndex = 39;
            txtJmFilter.Text = "Search Filter";
            txtJmFilter.Click += txtJmFilter_Click;
            txtJmFilter.TextChanged += txtJmFilter_TextChanged;
            txtJmFilter.Leave += txtJmFilter_Leave;
            // 
            // chkJm
            // 
            chkJm.AutoSize = true;
            chkJm.Cursor = Cursors.Hand;
            chkJm.Font = new Font("Segoe UI", 9.75F, FontStyle.Bold);
            chkJm.Location = new Point(300, 62);
            chkJm.Margin = new Padding(4, 5, 4, 5);
            chkJm.Name = "chkJm";
            chkJm.Size = new Size(127, 32);
            chkJm.TabIndex = 37;
            chkJm.Text = "Select All";
            chkJm.UseVisualStyleBackColor = true;
            chkJm.CheckedChanged += chkJm_CheckedChanged;
            // 
            // chkListJm
            // 
            chkListJm.CheckOnClick = true;
            chkListJm.Font = new Font("Consolas", 11F);
            chkListJm.FormattingEnabled = true;
            chkListJm.Location = new Point(9, 113);
            chkListJm.Margin = new Padding(4, 5, 4, 5);
            chkListJm.Name = "chkListJm";
            chkListJm.ScrollAlwaysVisible = true;
            chkListJm.Size = new Size(410, 214);
            chkListJm.TabIndex = 38;
            chkListJm.ItemCheck += chkListJm_ItemCheck;
            //chkListJm.SelectedIndexChanged += chkListJm_SelectedIndexChanged;
            // 
            // ComponentBreak
            // 
            AutoScaleDimensions = new SizeF(10F, 25F);
            AutoScaleMode = AutoScaleMode.Font;
            Controls.Add(groupBox3);
            Controls.Add(groupBox2);
            Controls.Add(groupBox4);
            Controls.Add(groupBox1);
            Margin = new Padding(4, 5, 4, 5);
            Name = "ComponentBreak";
            Size = new Size(893, 833);
            groupBox1.ResumeLayout(false);
            groupBox1.PerformLayout();
            groupBox2.ResumeLayout(false);
            groupBox2.PerformLayout();
            groupBox3.ResumeLayout(false);
            groupBox3.PerformLayout();
            groupBox4.ResumeLayout(false);
            groupBox4.PerformLayout();
            ResumeLayout(false);
        }

        #endregion
        private CheckedListBox chkListEqup;
        private GroupBox groupBox1;
        private GroupBox groupBox2;
        private TextBox txtBrkFilter;
        private CheckBox chkBrk;
        private CheckedListBox chkListBrk;
        private GroupBox groupBox3;
        private TextBox txtMiscFilter;
        private CheckBox chkMisc;
        private CheckedListBox chkListMisc;
        private GroupBox groupBox4;
        private TextBox txtJmFilter;
        private CheckBox chkJm;
        private CheckedListBox chkListJm;
        public CheckBox chkEqup;
        public TextBox txtEqupFilter;
        public Label lblEqupCount;
        public Label lblBrkCount;
        public Label lblMiscCount;
        public Label lblJmCount;
        private Label lblTotalEque;
        private Label lblTotalBreckcon;
        private Label lblTotalmis;
        private Label lblTotaljun;
    }
}
