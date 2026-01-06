using System.Windows.Forms;

namespace Electre_Customize_DotNet.ReportUI
{
    partial class CableListWindow
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
            groupBox3 = new GroupBox();
            lbltotalsheet = new Label();
            lblSheetCount = new Label();
            txtSheetFilter = new TextBox();
            chkSheet = new CheckBox();
            chkListSheet = new CheckedListBox();
            groupBox4 = new GroupBox();
            lbltotalLoom = new Label();
            chkLoom = new CheckBox();
            lblLoomCount = new Label();
            txtLoomFilter = new TextBox();
            chkListLoom = new CheckedListBox();
            chkProjSpec = new CheckBox();
            panel2 = new Panel();
            flowLayoutPanel1 = new FlowLayoutPanel();
            button1 = new Button();
            groupBox3.SuspendLayout();
            groupBox4.SuspendLayout();
            SuspendLayout();
            // 
            // groupBox3
            // 
            groupBox3.Controls.Add(lbltotalsheet);
            groupBox3.Controls.Add(lblSheetCount);
            groupBox3.Controls.Add(txtSheetFilter);
            groupBox3.Controls.Add(chkSheet);
            groupBox3.Controls.Add(chkListSheet);
            groupBox3.Font = new Font("Segoe UI", 12F, FontStyle.Bold);
            groupBox3.Location = new Point(457, 217);
            groupBox3.Margin = new Padding(7, 8, 7, 8);
            groupBox3.Name = "groupBox3";
            groupBox3.Padding = new Padding(4, 5, 4, 5);
            groupBox3.Size = new Size(429, 407);
            groupBox3.TabIndex = 45;
            groupBox3.TabStop = false;
            groupBox3.Text = "Sheet";
            // 
            // lbltotalsheet
            // 
            lbltotalsheet.AutoSize = true;
            lbltotalsheet.Font = new Font("Segoe UI Semibold", 9F, FontStyle.Bold | FontStyle.Italic, GraphicsUnit.Point, 0);
            lbltotalsheet.ForeColor = Color.Blue;
            lbltotalsheet.Location = new Point(266, 362);
            lbltotalsheet.Name = "lbltotalsheet";
            lbltotalsheet.Size = new Size(130, 25);
            lbltotalsheet.TabIndex = 42;
            lbltotalsheet.Text = "Total Looms: 0";
            // 
            // lblSheetCount
            // 
            lblSheetCount.AutoSize = true;
            lblSheetCount.Font = new Font("Segoe UI Semibold", 9F, FontStyle.Bold | FontStyle.Italic, GraphicsUnit.Point, 0);
            lblSheetCount.ForeColor = Color.Blue;
            lblSheetCount.Location = new Point(199, 335);
            lblSheetCount.Margin = new Padding(4, 0, 4, 0);
            lblSheetCount.Name = "lblSheetCount";
            lblSheetCount.Size = new Size(197, 25);
            lblSheetCount.TabIndex = 41;
            lblSheetCount.Text = "Selected Items Count: 0";
            // 
            // txtSheetFilter
            // 
            txtSheetFilter.Font = new Font("Segoe UI", 12F, FontStyle.Italic);
            txtSheetFilter.ForeColor = Color.Silver;
            txtSheetFilter.Location = new Point(9, 55);
            txtSheetFilter.Margin = new Padding(4, 5, 4, 5);
            txtSheetFilter.Name = "txtSheetFilter";
            txtSheetFilter.Size = new Size(278, 39);
            txtSheetFilter.TabIndex = 39;
            txtSheetFilter.Text = "Search Filter";
            txtSheetFilter.Click += txtSheetFilter_Click;
            txtSheetFilter.TextChanged += txtSheetFilter_TextChanged;
            txtSheetFilter.Leave += txtSheetFilter_Leave;
            // 
            // chkSheet
            // 
            chkSheet.AutoSize = true;
            chkSheet.Cursor = Cursors.Hand;
            chkSheet.Font = new Font("Segoe UI", 9.75F, FontStyle.Bold);
            chkSheet.Location = new Point(300, 62);
            chkSheet.Margin = new Padding(4, 5, 4, 5);
            chkSheet.Name = "chkSheet";
            chkSheet.Size = new Size(127, 32);
            chkSheet.TabIndex = 37;
            chkSheet.Text = "Select All";
            chkSheet.UseVisualStyleBackColor = true;
            chkSheet.CheckedChanged += chkSheet_CheckedChanged;
            // 
            // chkListSheet
            // 
            chkListSheet.CheckOnClick = true;
            chkListSheet.Font = new Font("Segoe UI", 11F);
            chkListSheet.FormattingEnabled = true;
            chkListSheet.Location = new Point(9, 113);
            chkListSheet.Margin = new Padding(4, 5, 4, 5);
            chkListSheet.Name = "chkListSheet";
            chkListSheet.ScrollAlwaysVisible = true;
            chkListSheet.Size = new Size(410, 208);
            chkListSheet.TabIndex = 38;
            chkListSheet.ItemCheck += chkListSheet_ItemCheck;
            // 
            // groupBox4
            // 
            groupBox4.Controls.Add(lbltotalLoom);
            groupBox4.Controls.Add(chkLoom);
            groupBox4.Controls.Add(lblLoomCount);
            groupBox4.Controls.Add(txtLoomFilter);
            groupBox4.Controls.Add(chkListLoom);
            groupBox4.Font = new Font("Segoe UI", 12F, FontStyle.Bold);
            groupBox4.Location = new Point(7, 217);
            groupBox4.Margin = new Padding(7, 8, 7, 8);
            groupBox4.Name = "groupBox4";
            groupBox4.Padding = new Padding(4, 5, 4, 5);
            groupBox4.Size = new Size(429, 407);
            groupBox4.TabIndex = 44;
            groupBox4.TabStop = false;
            groupBox4.Text = "Loom";
            // 
            // lbltotalLoom
            // 
            lbltotalLoom.AutoSize = true;
            lbltotalLoom.Font = new Font("Segoe UI Semibold", 9F, FontStyle.Bold | FontStyle.Italic, GraphicsUnit.Point, 0);
            lbltotalLoom.ForeColor = Color.Blue;
            lbltotalLoom.Location = new Point(271, 372);
            lbltotalLoom.Name = "lbltotalLoom";
            lbltotalLoom.Size = new Size(130, 25);
            lbltotalLoom.TabIndex = 42;
            lbltotalLoom.Text = "Total Looms: 0";
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
            // lblLoomCount
            // 
            lblLoomCount.AutoSize = true;
            lblLoomCount.Font = new Font("Segoe UI Semibold", 9F, FontStyle.Bold | FontStyle.Italic, GraphicsUnit.Point, 0);
            lblLoomCount.ForeColor = Color.Blue;
            lblLoomCount.Location = new Point(205, 342);
            lblLoomCount.Margin = new Padding(4, 0, 4, 0);
            lblLoomCount.Name = "lblLoomCount";
            lblLoomCount.Size = new Size(197, 25);
            lblLoomCount.TabIndex = 40;
            lblLoomCount.Text = "Selected Items Count: 0";
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
            // chkListLoom
            // 
            chkListLoom.CheckOnClick = true;
            chkListLoom.Font = new Font("Segoe UI", 11F);
            chkListLoom.FormattingEnabled = true;
            chkListLoom.Location = new Point(9, 113);
            chkListLoom.Margin = new Padding(4, 5, 4, 5);
            chkListLoom.Name = "chkListLoom";
            chkListLoom.ScrollAlwaysVisible = true;
            chkListLoom.Size = new Size(410, 208);
            chkListLoom.TabIndex = 38;
            chkListLoom.ItemCheck += chkListLoom_ItemCheck;
            // 
            // chkProjSpec
            // 
            chkProjSpec.AutoSize = true;
            chkProjSpec.Cursor = Cursors.Hand;
            chkProjSpec.Font = new Font("Segoe UI", 12F, FontStyle.Bold, GraphicsUnit.Point, 0);
            chkProjSpec.Location = new Point(16, 102);
            chkProjSpec.Margin = new Padding(4, 5, 4, 5);
            chkProjSpec.Name = "chkProjSpec";
            chkProjSpec.Size = new Size(237, 36);
            chkProjSpec.TabIndex = 40;
            chkProjSpec.Text = "Project Cable List";
            chkProjSpec.UseVisualStyleBackColor = true;
            // 
            // panel2
            // 
            panel2.Location = new Point(0, 0);
            panel2.Name = "panel2";
            panel2.Size = new Size(200, 100);
            panel2.TabIndex = 0;
            // 
            // flowLayoutPanel1
            // 
            flowLayoutPanel1.Location = new Point(0, 0);
            flowLayoutPanel1.Name = "flowLayoutPanel1";
            flowLayoutPanel1.Size = new Size(200, 100);
            flowLayoutPanel1.TabIndex = 0;
            // 
            // button1
            // 
            button1.Font = new Font("Segoe UI", 16F, FontStyle.Bold, GraphicsUnit.Point, 0);
            button1.ForeColor = Color.DeepSkyBlue;
            button1.Location = new Point(588, 65);
            button1.Name = "button1";
            button1.Size = new Size(265, 73);
            button1.TabIndex = 46;
            button1.Text = "Export to PDF";
            button1.UseVisualStyleBackColor = true;
            button1.Click += exportpdf_Click;
            // 
            // CableListWindow
            // 
            AutoScaleDimensions = new SizeF(10F, 25F);
            AutoScaleMode = AutoScaleMode.Font;
            Controls.Add(button1);
            Controls.Add(chkProjSpec);
            Controls.Add(groupBox3);
            Controls.Add(groupBox4);
            Margin = new Padding(7, 8, 7, 8);
            Name = "CableListWindow";
            Size = new Size(893, 640);
            groupBox3.ResumeLayout(false);
            groupBox3.PerformLayout();
            groupBox4.ResumeLayout(false);
            groupBox4.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private GroupBox groupBox3;
        private GroupBox groupBox4;
        public CheckedListBox chkListLoom;
        public Label lblSheetCount;
        public TextBox txtLoomFilter;
        public CheckBox chkLoom;
        public CheckBox chkProjSpec;
        public Label lblLoomCount;
        public TextBox txtSheetFilter;
        public CheckBox chkSheet;
        public CheckedListBox chkListSheet;
        public FlowLayoutPanel flowLayoutPanel1;
        public Panel panel2;
        private Label lbltotalLoom;
        private Label lbltotalsheet;
        private Button button1;
    }
}