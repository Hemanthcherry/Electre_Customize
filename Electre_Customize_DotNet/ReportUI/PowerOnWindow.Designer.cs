namespace Electre_Customize_DotNet.ReportUI
{
    partial class PowerOnWindow
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
            lblTotalPanel = new Label();
            lblPanelCount = new Label();
            txtPanelFilter = new TextBox();
            chkPanel = new CheckBox();
            chkListPanel = new CheckedListBox();
            chkProjSpec = new CheckBox();
            groupBox4.SuspendLayout();
            SuspendLayout();
            // 
            // groupBox4
            // 
            groupBox4.Controls.Add(lblTotalPanel);
            groupBox4.Controls.Add(lblPanelCount);
            groupBox4.Controls.Add(txtPanelFilter);
            groupBox4.Controls.Add(chkPanel);
            groupBox4.Controls.Add(chkListPanel);
            groupBox4.Font = new Font("Segoe UI", 12F, FontStyle.Bold);
            groupBox4.Location = new Point(197, 129);
            groupBox4.Margin = new Padding(7, 8, 7, 8);
            groupBox4.Name = "groupBox4";
            groupBox4.Padding = new Padding(4, 5, 4, 5);
            groupBox4.Size = new Size(429, 445);
            groupBox4.TabIndex = 45;
            groupBox4.TabStop = false;
            groupBox4.Text = "Panel";
            // 
            // lblTotalPanel
            // 
            lblTotalPanel.AutoSize = true;
            lblTotalPanel.Font = new Font("Segoe UI Semibold", 9F, FontStyle.Bold | FontStyle.Italic, GraphicsUnit.Point, 0);
            lblTotalPanel.ForeColor = Color.Blue;
            lblTotalPanel.Location = new Point(219, 411);
            lblTotalPanel.Name = "lblTotalPanel";
            lblTotalPanel.Size = new Size(130, 25);
            lblTotalPanel.TabIndex = 42;
            lblTotalPanel.Text = "Total Panels: 0";
            // 
            // lblPanelCount
            // 
            lblPanelCount.AutoSize = true;
            lblPanelCount.Font = new Font("Segoe UI Semibold", 9F, FontStyle.Bold | FontStyle.Italic, GraphicsUnit.Point, 0);
            lblPanelCount.ForeColor = Color.Blue;
            lblPanelCount.Location = new Point(223, 379);
            lblPanelCount.Margin = new Padding(4, 0, 4, 0);
            lblPanelCount.Name = "lblPanelCount";
            lblPanelCount.Size = new Size(189, 25);
            lblPanelCount.TabIndex = 40;
            lblPanelCount.Text = "Selected Item Count: 0";
            // 
            // txtPanelFilter
            // 
            txtPanelFilter.Font = new Font("Segoe UI", 12F, FontStyle.Italic);
            txtPanelFilter.ForeColor = Color.Silver;
            txtPanelFilter.Location = new Point(9, 55);
            txtPanelFilter.Margin = new Padding(4, 5, 4, 5);
            txtPanelFilter.Name = "txtPanelFilter";
            txtPanelFilter.Size = new Size(278, 39);
            txtPanelFilter.TabIndex = 39;
            txtPanelFilter.Text = "Search Filter";
            txtPanelFilter.Click += txtPanelFilter_Click;
            txtPanelFilter.TextChanged += txtPanelFilter_TextChanged;
            txtPanelFilter.Leave += txtPanelFilter_Leave;
            // 
            // chkPanel
            // 
            chkPanel.AutoSize = true;
            chkPanel.Cursor = Cursors.Hand;
            chkPanel.Font = new Font("Segoe UI", 9.75F, FontStyle.Bold);
            chkPanel.Location = new Point(300, 62);
            chkPanel.Margin = new Padding(4, 5, 4, 5);
            chkPanel.Name = "chkPanel";
            chkPanel.Size = new Size(127, 32);
            chkPanel.TabIndex = 37;
            chkPanel.Text = "Select All";
            chkPanel.UseVisualStyleBackColor = true;
            chkPanel.CheckedChanged += chkPanel_CheckedChanged;
            // 
            // chkListPanel
            // 
            chkListPanel.CheckOnClick = true;
            chkListPanel.Font = new Font("Segoe UI", 9F);
            chkListPanel.FormattingEnabled = true;
            chkListPanel.Location = new Point(9, 113);
            chkListPanel.Margin = new Padding(4, 5, 4, 5);
            chkListPanel.Name = "chkListPanel";
            chkListPanel.ScrollAlwaysVisible = true;
            chkListPanel.Size = new Size(410, 256);
            chkListPanel.TabIndex = 38;
            chkListPanel.ItemCheck += chkListPanel_ItemCheck;
            // 
            // chkProjSpec
            // 
            chkProjSpec.AutoSize = true;
            chkProjSpec.Cursor = Cursors.Hand;
            chkProjSpec.Font = new Font("Segoe UI", 12F, FontStyle.Bold, GraphicsUnit.Point, 0);
            chkProjSpec.Location = new Point(197, 40);
            chkProjSpec.Margin = new Padding(4, 5, 4, 5);
            chkProjSpec.Name = "chkProjSpec";
            chkProjSpec.Size = new Size(216, 36);
            chkProjSpec.TabIndex = 46;
            chkProjSpec.Text = "Project Specific";
            chkProjSpec.UseVisualStyleBackColor = true;
            // 
            // PowerOnWindow
            // 
            AutoScaleDimensions = new SizeF(10F, 25F);
            AutoScaleMode = AutoScaleMode.Font;
            Controls.Add(chkProjSpec);
            Controls.Add(groupBox4);
            Margin = new Padding(7, 8, 7, 8);
            Name = "PowerOnWindow";
            Size = new Size(893, 640);
            groupBox4.ResumeLayout(false);
            groupBox4.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private GroupBox groupBox4;
        public Label lblPanelCount;
        public TextBox txtPanelFilter;
        public CheckBox chkPanel;
        public CheckedListBox chkListPanel;



        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        /*   private void InitializeComponent()
           {
               components = new System.ComponentModel.Container();
               this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
           }*/

        #endregion

        public CheckBox chkProjSpec;
        private Label lblTotalPanel;
    }
}
