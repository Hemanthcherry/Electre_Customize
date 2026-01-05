namespace Electre_Customize_DotNet.ReportUI
{
    partial class CustomMessageBox
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null)) components.Dispose();
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            lblMessage = new Label();
            btnWithConfig = new Button();
            btnWithoutConfig = new Button();
            SuspendLayout();
            // 
            // lblMessage
            // 
            lblMessage.AutoSize = true;
            lblMessage.Location = new Point(20, 20);
            lblMessage.Name = "lblMessage";
            lblMessage.Size = new Size(82, 25);
            lblMessage.TabIndex = 0;
            lblMessage.Text = "Message";
            // 
            // btnWithConfig
            // 
            btnWithConfig.BackColor = SystemColors.Highlight;
            btnWithConfig.Font = new Font("Segoe UI", 10F, FontStyle.Bold, GraphicsUnit.Point, 0);
            btnWithConfig.ForeColor = SystemColors.ButtonHighlight;
            btnWithConfig.Location = new Point(40, 69);
            btnWithConfig.Name = "btnWithConfig";
            btnWithConfig.Size = new Size(161, 43);
            btnWithConfig.TabIndex = 1;
            btnWithConfig.Text = "With Config";
            btnWithConfig.UseVisualStyleBackColor = false;
            btnWithConfig.Click += btnWithConfig_Click;
            // 
            // btnWithoutConfig
            // 
            btnWithoutConfig.BackColor = SystemColors.Highlight;
            btnWithoutConfig.Font = new Font("Segoe UI", 10F, FontStyle.Bold, GraphicsUnit.Point, 0);
            btnWithoutConfig.ForeColor = SystemColors.ButtonHighlight;
            btnWithoutConfig.Location = new Point(230, 69);
            btnWithoutConfig.Name = "btnWithoutConfig";
            btnWithoutConfig.Size = new Size(169, 43);
            btnWithoutConfig.TabIndex = 2;
            btnWithoutConfig.Text = "Without Config";
            btnWithoutConfig.UseVisualStyleBackColor = false;
            btnWithoutConfig.Click += btnWithoutConfig_Click;
            // 
            // CustomMessageBox
            // 
            ClientSize = new Size(475, 148);
            Controls.Add(btnWithoutConfig);
            Controls.Add(btnWithConfig);
            Controls.Add(lblMessage);
            Name = "CustomMessageBox";
            StartPosition = FormStartPosition.CenterParent;
            ResumeLayout(false);
            PerformLayout();
        }

        private System.Windows.Forms.Label lblMessage;
        private System.Windows.Forms.Button btnWithConfig;
        private System.Windows.Forms.Button btnWithoutConfig;
    }
}