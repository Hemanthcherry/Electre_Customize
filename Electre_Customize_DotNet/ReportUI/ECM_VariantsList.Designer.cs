namespace Electre_Customize_DotNet.ReportUI
{
    partial class ECM_VariantsList
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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            grpBoxVariants = new GroupBox();
            lblTotalVariants = new Label();
            txtVariantFilter = new TextBox();
            chkListVariants = new CheckedListBox();
            btnOK = new Button();
            grpBoxVariants.SuspendLayout();
            SuspendLayout();
            // 
            // grpBoxVariants
            // 
            grpBoxVariants.Controls.Add(lblTotalVariants);
            grpBoxVariants.Controls.Add(txtVariantFilter);
            grpBoxVariants.Controls.Add(chkListVariants);
            grpBoxVariants.Font = new Font("Segoe UI", 12F, FontStyle.Bold);
            grpBoxVariants.Location = new Point(64, 17);
            grpBoxVariants.Margin = new Padding(7, 8, 7, 8);
            grpBoxVariants.Name = "grpBoxVariants";
            grpBoxVariants.Padding = new Padding(4, 5, 4, 5);
            grpBoxVariants.Size = new Size(429, 417);
            grpBoxVariants.TabIndex = 46;
            grpBoxVariants.TabStop = false;
            grpBoxVariants.Text = "Variants";
            // 
            // lblTotalVariants
            // 
            lblTotalVariants.AutoSize = true;
            lblTotalVariants.Font = new Font("Segoe UI Semibold", 9F, FontStyle.Bold | FontStyle.Italic, GraphicsUnit.Point, 0);
            lblTotalVariants.ForeColor = Color.Blue;
            lblTotalVariants.Location = new Point(278, 380);
            lblTotalVariants.Name = "lblTotalVariants";
            lblTotalVariants.Size = new Size(144, 25);
            lblTotalVariants.TabIndex = 41;
            lblTotalVariants.Text = "Total Varients: 0";
            // 
            // txtVariantFilter
            // 
            txtVariantFilter.Font = new Font("Segoe UI", 12F, FontStyle.Italic);
            txtVariantFilter.ForeColor = Color.Silver;
            txtVariantFilter.Location = new Point(9, 55);
            txtVariantFilter.Margin = new Padding(4, 5, 4, 5);
            txtVariantFilter.Name = "txtVariantFilter";
            txtVariantFilter.Size = new Size(278, 39);
            txtVariantFilter.TabIndex = 39;
            txtVariantFilter.Text = "Search Filter";
            txtVariantFilter.Click += txtVariantFilter_Click;
            txtVariantFilter.TextChanged += txtVariantFilter_TextChanged;
            txtVariantFilter.Leave += txtVariantFilter_Leave;
            // 
            // chkListVariants
            // 
            chkListVariants.CheckOnClick = true;
            chkListVariants.Font = new Font("Segoe UI", 11F);
            chkListVariants.FormattingEnabled = true;
            chkListVariants.Location = new Point(9, 113);
            chkListVariants.Margin = new Padding(4, 5, 4, 5);
            chkListVariants.Name = "chkListVariants";
            chkListVariants.ScrollAlwaysVisible = true;
            chkListVariants.Size = new Size(410, 242);
            chkListVariants.TabIndex = 38;
            //chkListVariants.ItemCheck += chkListLoom_ItemCheck;

            // 
            // btnOK
            // 
            btnOK.BackColor = SystemColors.Highlight;
            btnOK.ForeColor = SystemColors.ControlLightLight;
            btnOK.Location = new Point(194, 445);
            btnOK.Name = "btnOK";
            btnOK.Size = new Size(112, 41);
            btnOK.TabIndex = 47;
            btnOK.Text = "OK";
            btnOK.UseVisualStyleBackColor = false;
            btnOK.Click += btnOK_Click;
            // 
            // ECM_VariantsList
            // 
            AutoScaleDimensions = new SizeF(10F, 25F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(572, 499);
            Controls.Add(btnOK);
            Controls.Add(grpBoxVariants);
            Name = "ECM_VariantsList";
            Text = "ECM_VariantsList";
            grpBoxVariants.ResumeLayout(false);
            grpBoxVariants.PerformLayout();
            ResumeLayout(false);
        }

        #endregion

        private GroupBox grpBoxVariants;
        private Label lblTotalVariants;
        public TextBox txtVariantFilter;
        public CheckedListBox chkListVariants;
        private Button btnOK;
    }
}