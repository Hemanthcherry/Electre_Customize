namespace Electre_Customize_DotNet.ReportUI
{
    partial class ContinuityWindow
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
            groupBox4 = new GroupBox();
            lblTotalLoom = new Label();
            lblLoomCount = new Label();
            txtLoomFilter = new TextBox();
            chkLoom = new CheckBox();
            chkListLoom = new CheckedListBox();
            groupBox4.SuspendLayout();
            SuspendLayout();
            // 
            // groupBox4
            // 
            groupBox4.Controls.Add(lblTotalLoom);
            groupBox4.Controls.Add(lblLoomCount);
            groupBox4.Controls.Add(txtLoomFilter);
            groupBox4.Controls.Add(chkLoom);
            groupBox4.Controls.Add(chkListLoom);
            groupBox4.Font = new Font("Segoe UI", 12F, FontStyle.Bold);
            groupBox4.Location = new Point(231, 165);
            groupBox4.Margin = new Padding(7, 8, 7, 8);
            groupBox4.Name = "groupBox4";
            groupBox4.Padding = new Padding(4, 5, 4, 5);
            groupBox4.Size = new Size(429, 505);
            groupBox4.TabIndex = 45;
            groupBox4.TabStop = false;
            groupBox4.Text = "Loom";
           // groupBox4.Enter += groupBox4_Enter;
            // 
            // lblTotalLoom
            // 
            lblTotalLoom.AutoSize = true;
            lblTotalLoom.Font = new Font("Segoe UI Semibold", 9F, FontStyle.Bold | FontStyle.Italic, GraphicsUnit.Point, 0);
            lblTotalLoom.ForeColor = Color.Blue;
            lblTotalLoom.Location = new Point(275, 456);
            lblTotalLoom.Name = "lblTotalLoom";
            lblTotalLoom.Size = new Size(111, 25);
            lblTotalLoom.TabIndex = 41;
            lblTotalLoom.Text = "Total Looms: 0";
            // 
            // lblLoomCount
            // 
            lblLoomCount.AutoSize = true;
            lblLoomCount.Font = new Font("Segoe UI Semibold", 9F, FontStyle.Bold | FontStyle.Italic, GraphicsUnit.Point, 0);
            lblLoomCount.ForeColor = Color.Blue;
            lblLoomCount.Location = new Point(207, 428);
            lblLoomCount.Margin = new Padding(4, 0, 4, 0);
            lblLoomCount.Name = "lblLoomCount";
            lblLoomCount.Size = new Size(197, 25);
            lblLoomCount.TabIndex = 40;
            lblLoomCount.Text = "Selected Items Count: 0";
           // lblLoomCount.Click += lblLoomCount_Click;
            // 
            // txtLoomFilter
            // 
            txtLoomFilter.Font = new Font("Segoe UI", 12F, FontStyle.Italic);
            txtLoomFilter.ForeColor = Color.Silver;
            txtLoomFilter.Location = new Point(9, 55);
            txtLoomFilter.Margin = new Padding(4, 5, 4, 5);
            txtLoomFilter.Name = "txtLoomFilter";
            txtLoomFilter.Size = new Size(278, 39);
            txtLoomFilter.TabIndex = 39;
            txtLoomFilter.Text = "Search Filter";
            txtLoomFilter.Click += txtLoomFilter_Click;
            txtLoomFilter.TextChanged += txtLoomFilter_TextChanged;
            txtLoomFilter.Leave += txtLoomFilter_Leave;
            // 
            // chkLoom
            // 
            chkLoom.AutoSize = true;
            chkLoom.Cursor = Cursors.Hand;
            chkLoom.Font = new Font("Segoe UI", 9.75F, FontStyle.Bold);
            chkLoom.Location = new Point(300, 62);
            chkLoom.Margin = new Padding(4, 5, 4, 5);
            chkLoom.Name = "chkLoom";
            chkLoom.Size = new Size(127, 32);
            chkLoom.TabIndex = 37;
            chkLoom.Text = "Select All";
            chkLoom.UseVisualStyleBackColor = true;
            chkLoom.CheckedChanged += chkLoom_CheckedChanged;
            // 
            // chkListLoom
            // 
            chkListLoom.CheckOnClick = true;
            chkListLoom.Font = new Font("Segoe UI", 11F);
            chkListLoom.FormattingEnabled = true;
            chkListLoom.Location = new Point(9, 113);
            chkListLoom.Margin = new Padding(4, 5, 4, 5);
            chkListLoom.Name = "chkListLoom";
            chkListLoom.ScrollAlwaysVisible = true;
            chkListLoom.Size = new Size(410, 310);
            chkListLoom.TabIndex = 38;
            chkListLoom.ItemCheck += chkListLoom_ItemCheck;
           // chkListLoom.SelectedIndexChanged += chkListLoom_SelectedIndexChanged;
            // 
            // ContinuityWindow
            // 
            AutoScaleDimensions = new SizeF(10F, 25F);
            AutoScaleMode = AutoScaleMode.Font;
            Controls.Add(groupBox4);
            Margin = new Padding(7, 8, 7, 8);
            Name = "ContinuityWindow";
            Size = new Size(893, 833);
            groupBox4.ResumeLayout(false);
            groupBox4.PerformLayout();
            ResumeLayout(false);
        }

        #endregion

        private GroupBox groupBox4;
        public Label lblLoomCount;
        public TextBox txtLoomFilter;
        public CheckBox chkLoom;
        public CheckedListBox chkListLoom;
        private Label lblTotalLoom;
    }
}
