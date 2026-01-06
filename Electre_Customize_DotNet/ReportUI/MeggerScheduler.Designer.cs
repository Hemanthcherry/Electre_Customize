namespace Electre_Customize_DotNet.ReportUI
{
    partial class MeggerScheduler
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
            groupBox1 = new GroupBox();
            labelTotal = new Label();
            lblLoomCount = new Label();
            chkListLoom = new CheckedListBox();
            chkLoom = new CheckBox();
            txtLoomFilter = new TextBox();
            groupBox1.SuspendLayout();
            SuspendLayout();
            // 
            // groupBox1
            // 
            groupBox1.Controls.Add(labelTotal);
            groupBox1.Controls.Add(lblLoomCount);
            groupBox1.Controls.Add(chkListLoom);
            groupBox1.Controls.Add(chkLoom);
            groupBox1.Controls.Add(txtLoomFilter);
            groupBox1.Font = new Font("Segoe UI", 12F, FontStyle.Bold);
            groupBox1.Location = new Point(215, 64);
            groupBox1.Margin = new Padding(7, 8, 7, 8);
            groupBox1.Name = "groupBox1";
            groupBox1.Padding = new Padding(4, 5, 4, 5);
            groupBox1.Size = new Size(429, 468);
            groupBox1.TabIndex = 45;
            groupBox1.TabStop = false;
            groupBox1.Text = "Loom";
            // 
            // labelTotal
            // 
            labelTotal.AutoSize = true;
            labelTotal.Font = new Font("Segoe UI Semibold", 9F, FontStyle.Bold | FontStyle.Italic, GraphicsUnit.Point, 0);
            labelTotal.ForeColor = Color.Blue;
            labelTotal.Location = new Point(269, 433);
            labelTotal.Name = "labelTotal";
            labelTotal.Size = new Size(130, 25);
            labelTotal.TabIndex = 41;
            labelTotal.Text = "Total Looms: 0";
            // 
            // lblLoomCount
            // 
            lblLoomCount.AutoSize = true;
            lblLoomCount.Font = new Font("Segoe UI Semibold", 9F, FontStyle.Bold | FontStyle.Italic, GraphicsUnit.Point, 0);
            lblLoomCount.ForeColor = Color.Blue;
            lblLoomCount.Location = new Point(202, 408);
            lblLoomCount.Margin = new Padding(4, 0, 4, 0);
            lblLoomCount.Name = "lblLoomCount";
            lblLoomCount.Size = new Size(197, 25);
            lblLoomCount.TabIndex = 40;
            lblLoomCount.Text = "Selected Items Count: 0";
            // 
            // chkListLoom
            // 
            chkListLoom.CheckOnClick = true;
            chkListLoom.Font = new Font("Segoe UI", 11F);
            chkListLoom.FormattingEnabled = true;
            chkListLoom.Location = new Point(7, 127);
            chkListLoom.Margin = new Padding(4, 5, 4, 5);
            chkListLoom.Name = "chkListLoom";
            chkListLoom.ScrollAlwaysVisible = true;
            chkListLoom.Size = new Size(410, 276);
            chkListLoom.TabIndex = 38;
            chkListLoom.ItemCheck += checkedListLoom_ItemCheck;
            chkListLoom.SelectedIndexChanged += checkedListLoom_SelectedIndexChanged;
            // 
            // chkLoom
            // 
            chkLoom.AutoSize = true;
            chkLoom.Cursor = Cursors.Hand;
            chkLoom.Font = new Font("Segoe UI", 9.75F, FontStyle.Bold);
            chkLoom.Location = new Point(293, 65);
            chkLoom.Margin = new Padding(4, 5, 4, 5);
            chkLoom.Name = "chkLoom";
            chkLoom.Size = new Size(127, 32);
            chkLoom.TabIndex = 37;
            chkLoom.Text = "Select All";
            chkLoom.UseVisualStyleBackColor = true;
            chkLoom.CheckedChanged += checkLoom_CheckedChanged;
            // 
            // txtLoomFilter
            // 
            txtLoomFilter.Font = new Font("Segoe UI", 12F, FontStyle.Italic);
            txtLoomFilter.ForeColor = Color.Silver;
            txtLoomFilter.Location = new Point(7, 58);
            txtLoomFilter.Margin = new Padding(4, 5, 4, 5);
            txtLoomFilter.Name = "txtLoomFilter";
            txtLoomFilter.Size = new Size(278, 39);
            txtLoomFilter.TabIndex = 0;
            txtLoomFilter.Text = "Search Filter";
            txtLoomFilter.Click += textLoomFilter_Click;
            txtLoomFilter.TextChanged += textLoomFilter_TextChanged;
            txtLoomFilter.Leave += textLoomFilter_Leave;
            // 
            // MeggerScheduler
            // 
            AutoScaleDimensions = new SizeF(10F, 25F);
            AutoScaleMode = AutoScaleMode.Font;
            Controls.Add(groupBox1);
            Name = "MeggerScheduler";
            Size = new Size(893, 640);
            groupBox1.ResumeLayout(false);
            groupBox1.PerformLayout();
            ResumeLayout(false);
        }

        #endregion

        private CheckBox meggerscheduler;
        private GroupBox groupBox1;
        private TextBox txtLoomFilter;
        private CheckBox chkLoom;
        private CheckedListBox chkListLoom;
        private Label lblLoomCount;
        private Label labelTotal;
    }
}
