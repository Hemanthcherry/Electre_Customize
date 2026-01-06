namespace Electre_Customize_DotNet.ReportUI
{
    partial class PanelDrawingWindow
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
            lblPanelCount = new Label();
            chkListPanel = new CheckedListBox();
            btnProceed = new Button();
            txtPanelDigit = new TextBox();
            panelExtractionDoc = new LinkLabel();
            groupBox1.SuspendLayout();
            SuspendLayout();
            // 
            // groupBox1
            // 
            groupBox1.Controls.Add(lblPanelCount);
            groupBox1.Controls.Add(chkListPanel);
            groupBox1.Controls.Add(btnProceed);
            groupBox1.Font = new Font("Segoe UI", 10F, FontStyle.Bold, GraphicsUnit.Point, 0);
            groupBox1.Location = new Point(210, 112);
            groupBox1.Name = "groupBox1";
            groupBox1.Size = new Size(485, 363);
            groupBox1.TabIndex = 5;
            groupBox1.TabStop = false;
            groupBox1.Text = "Panels";
            // 
            // lblPanelCount
            // 
            lblPanelCount.AutoSize = true;
            lblPanelCount.Font = new Font("Segoe UI Semibold", 9F, FontStyle.Bold | FontStyle.Italic, GraphicsUnit.Point, 0);
            lblPanelCount.ForeColor = Color.Blue;
            lblPanelCount.Location = new Point(51, 306);
            lblPanelCount.Margin = new Padding(4, 0, 4, 0);
            lblPanelCount.Name = "lblPanelCount";
            lblPanelCount.Size = new Size(189, 25);
            lblPanelCount.TabIndex = 41;
            lblPanelCount.Text = "Selected Item Count: 0";
            lblPanelCount.Visible = false;
            // 
            // chkListPanel
            // 
            chkListPanel.CheckOnClick = true;
            chkListPanel.Font = new Font("Segoe UI", 10F, FontStyle.Regular, GraphicsUnit.Point, 0);
            chkListPanel.FormattingEnabled = true;
            chkListPanel.Location = new Point(12, 47);
            chkListPanel.Name = "chkListPanel";
            chkListPanel.ScrollAlwaysVisible = true;
            chkListPanel.Size = new Size(450, 252);
            chkListPanel.TabIndex = 1;
            chkListPanel.ItemCheck += chkListPanel_ItemCheck;
            // 
            // btnProceed
            // 
            btnProceed.BackColor = SystemColors.GradientActiveCaption;
            btnProceed.Location = new Point(316, 312);
            btnProceed.Name = "btnProceed";
            btnProceed.Size = new Size(125, 42);
            btnProceed.TabIndex = 0;
            btnProceed.Text = "Proceed";
            btnProceed.UseVisualStyleBackColor = false;
            btnProceed.Click += btnProceed_Click;
            // 
            // txtPanelDigit
            // 
            txtPanelDigit.Location = new Point(467, 482);
            txtPanelDigit.Name = "txtPanelDigit";
            txtPanelDigit.PlaceholderText = "PanelDigit";
            txtPanelDigit.Size = new Size(150, 31);
            txtPanelDigit.TabIndex = 29;
            // 
            // panelExtractionDoc
            // 
            panelExtractionDoc.AutoSize = true;
            panelExtractionDoc.Font = new Font("Segoe UI", 11F, FontStyle.Bold, GraphicsUnit.Point, 0);
            panelExtractionDoc.Location = new Point(118, 42);
            panelExtractionDoc.Name = "panelExtractionDoc";
            panelExtractionDoc.Size = new Size(533, 30);
            panelExtractionDoc.TabIndex = 30;
            panelExtractionDoc.TabStop = true;
            panelExtractionDoc.Text = "Click Here for the Creation of Panel Drawing Steps";
            panelExtractionDoc.LinkClicked += linkLabel1_LinkClicked;
            // 
            // PanelDrawingWindow
            // 
            AutoScaleDimensions = new SizeF(10F, 25F);
            AutoScaleMode = AutoScaleMode.Font;
            Controls.Add(panelExtractionDoc);
            Controls.Add(txtPanelDigit);
            Controls.Add(groupBox1);
            Name = "PanelDrawingWindow";
            Size = new Size(893, 640);
            groupBox1.ResumeLayout(false);
            groupBox1.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion
        private GroupBox groupBox1;
        private Button btnProceed;
        private CheckedListBox chkListPanel;
        public Label lblPanelCount;
        private TextBox txtPanelDigit;
        private LinkLabel panelExtractionDoc;
    }
}
