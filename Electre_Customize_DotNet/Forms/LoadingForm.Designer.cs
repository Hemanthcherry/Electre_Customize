namespace Electre_Customize_DotNet
{
    partial class LoadingForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(LoadingForm));
            loadingbtn = new Button();
            lableLoadMessage = new Label();
            completionlbl = new Label();
            reportLocLinklbl = new LinkLabel();
            reportLoclbl = new Label();
            SuspendLayout();
            // 
            // loadingbtn
            // 
            loadingbtn.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            loadingbtn.Cursor = Cursors.Hand;
            loadingbtn.Font = new Font("Segoe UI Semibold", 12F, FontStyle.Bold, GraphicsUnit.Point, 0);
            loadingbtn.Location = new Point(454, 108);
            loadingbtn.Margin = new Padding(2);
            loadingbtn.Name = "loadingbtn";
            loadingbtn.Size = new Size(109, 35);
            loadingbtn.TabIndex = 0;
            loadingbtn.Text = "OK";
            loadingbtn.UseVisualStyleBackColor = true;
            loadingbtn.TextChanged += LoadingForm_Load;
            loadingbtn.Click += Loadingbtn_Click;
            // 
            // lableLoadMessage
            // 
            lableLoadMessage.AutoSize = true;
            lableLoadMessage.Font = new Font("Segoe UI Semibold", 12F, FontStyle.Bold, GraphicsUnit.Point, 0);
            lableLoadMessage.ForeColor = Color.Blue;
            lableLoadMessage.Location = new Point(8, 13);
            lableLoadMessage.Margin = new Padding(2, 0, 2, 0);
            lableLoadMessage.Name = "lableLoadMessage";
            lableLoadMessage.Size = new Size(0, 21);
            lableLoadMessage.TabIndex = 2;
            // 
            // completionlbl
            // 
            completionlbl.AutoSize = true;
            completionlbl.Font = new Font("Segoe UI Black", 18F, FontStyle.Bold, GraphicsUnit.Point, 0);
            completionlbl.ForeColor = Color.Blue;
            completionlbl.Location = new Point(41, 200);
            completionlbl.Margin = new Padding(2, 0, 2, 0);
            completionlbl.Name = "completionlbl";
            completionlbl.Size = new Size(0, 32);
            completionlbl.TabIndex = 4;
            // 
            // reportLocLinklbl
            // 
            reportLocLinklbl.AutoSize = true;
            reportLocLinklbl.Location = new Point(215, 257);
            reportLocLinklbl.Margin = new Padding(2, 0, 2, 0);
            reportLocLinklbl.Name = "reportLocLinklbl";
            reportLocLinklbl.Size = new Size(79, 15);
            reportLocLinklbl.TabIndex = 5;
            reportLocLinklbl.TabStop = true;
            reportLocLinklbl.Text = "Click to Open";
            reportLocLinklbl.LinkClicked += reportLocLinklbl_LinkClicked;
            // 
            // reportLoclbl
            // 
            reportLoclbl.AutoSize = true;
            reportLoclbl.Location = new Point(27, 239);
            reportLoclbl.Margin = new Padding(2, 0, 2, 0);
            reportLoclbl.Name = "reportLoclbl";
            reportLoclbl.Size = new Size(0, 15);
            reportLoclbl.TabIndex = 6;
            // 
            // LoadingForm
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(571, 281);
            Controls.Add(reportLoclbl);
            Controls.Add(reportLocLinklbl);
            Controls.Add(completionlbl);
            Controls.Add(lableLoadMessage);
            Controls.Add(loadingbtn);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            Icon = (Icon)resources.GetObject("$this.Icon");
            Margin = new Padding(2);
            MaximizeBox = false;
            Name = "LoadingForm";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "LoadingForm";
            Load += LoadingForm_Load;
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button loadingbtn;
        private Label lableLoadMessage;
        private Label completionlbl;
        private LinkLabel reportLocLinklbl;
        private Label reportLoclbl;
    }
}