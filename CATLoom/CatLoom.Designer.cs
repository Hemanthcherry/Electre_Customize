namespace CATLoom
{
    partial class CATLoomWindow
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(CATLoomWindow));
            catLoomTxt = new TextBox();
            catLoomBtn = new Button();
            catLoomSearchTxt = new TextBox();
            catSearchChkBox = new CheckBox();
            catFilelbl = new Label();
            pictureBox1 = new PictureBox();
            Browser = new Label();
            groupBox1 = new GroupBox();
            chkCatLoomList = new CheckedListBox();
            ((System.ComponentModel.ISupportInitialize)pictureBox1).BeginInit();
            groupBox1.SuspendLayout();
            SuspendLayout();
            // 
            // catLoomTxt
            // 
            catLoomTxt.Font = new Font("Segoe UI", 10F, FontStyle.Italic, GraphicsUnit.Point, 0);
            catLoomTxt.ForeColor = Color.FromArgb(192, 64, 0);
            catLoomTxt.Location = new Point(58, 68);
            catLoomTxt.Multiline = true;
            catLoomTxt.Name = "catLoomTxt";
            catLoomTxt.Size = new Size(728, 75);
            catLoomTxt.TabIndex = 0;
            catLoomTxt.TextChanged += catLoomTxt_TextChanged;
            // 
            // catLoomBtn
            // 
            catLoomBtn.AutoSize = true;
            catLoomBtn.Cursor = Cursors.Hand;
            catLoomBtn.Font = new Font("Segoe UI", 16F, FontStyle.Bold, GraphicsUnit.Point, 0);
            catLoomBtn.ForeColor = Color.Blue;
            catLoomBtn.Location = new Point(264, 451);
            catLoomBtn.Name = "catLoomBtn";
            catLoomBtn.Size = new Size(443, 58);
            catLoomBtn.TabIndex = 1;
            catLoomBtn.Text = "Generate LoomWise Report";
            catLoomBtn.UseVisualStyleBackColor = true;
            catLoomBtn.Click += catLoomBtn_Click;
            // 
            // catLoomSearchTxt
            // 
            catLoomSearchTxt.Font = new Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point, 0);
            catLoomSearchTxt.Location = new Point(6, 38);
            catLoomSearchTxt.Name = "catLoomSearchTxt";
            catLoomSearchTxt.Size = new Size(277, 31);
            catLoomSearchTxt.TabIndex = 2;
            catLoomSearchTxt.TextChanged += catLoomSearchTxt_TextChanged;
            // 
            // catSearchChkBox
            // 
            catSearchChkBox.AutoSize = true;
            catSearchChkBox.Cursor = Cursors.Hand;
            catSearchChkBox.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, 0);
            catSearchChkBox.ForeColor = Color.Blue;
            catSearchChkBox.Location = new Point(309, 38);
            catSearchChkBox.Name = "catSearchChkBox";
            catSearchChkBox.Size = new Size(117, 29);
            catSearchChkBox.TabIndex = 4;
            catSearchChkBox.Text = "Select All";
            catSearchChkBox.UseVisualStyleBackColor = true;
            catSearchChkBox.CheckedChanged += catSearchChkBox_CheckedChanged;
            // 
            // catFilelbl
            // 
            catFilelbl.AutoSize = true;
            catFilelbl.Font = new Font("Segoe UI Semibold", 12F, FontStyle.Bold, GraphicsUnit.Point, 0);
            catFilelbl.ForeColor = Color.Blue;
            catFilelbl.Location = new Point(62, 28);
            catFilelbl.Name = "catFilelbl";
            catFilelbl.Size = new Size(200, 32);
            catFilelbl.TabIndex = 5;
            catFilelbl.Text = "CAT File Location";
            catFilelbl.Click += label1_Click;
            // 
            // pictureBox1
            // 
            pictureBox1.Cursor = Cursors.Hand;
            pictureBox1.Image = (Image)resources.GetObject("pictureBox1.Image");
            pictureBox1.Location = new Point(792, 68);
            pictureBox1.Name = "pictureBox1";
            pictureBox1.Size = new Size(150, 75);
            pictureBox1.SizeMode = PictureBoxSizeMode.StretchImage;
            pictureBox1.TabIndex = 6;
            pictureBox1.TabStop = false;
            pictureBox1.Click += pictureBox1_Click;
            // 
            // Browser
            // 
            Browser.AutoSize = true;
            Browser.Font = new Font("Segoe UI Semibold", 12F, FontStyle.Bold, GraphicsUnit.Point, 0);
            Browser.ForeColor = Color.Blue;
            Browser.Location = new Point(824, 28);
            Browser.Name = "Browser";
            Browser.Size = new Size(93, 32);
            Browser.TabIndex = 7;
            Browser.Text = "Browse";
            Browser.Click += label1_Click;
            // 
            // groupBox1
            // 
            groupBox1.Controls.Add(chkCatLoomList);
            groupBox1.Controls.Add(catLoomSearchTxt);
            groupBox1.Controls.Add(catSearchChkBox);
            groupBox1.Font = new Font("Segoe UI", 12F, FontStyle.Bold, GraphicsUnit.Point, 0);
            groupBox1.ForeColor = Color.Blue;
            groupBox1.Location = new Point(264, 163);
            groupBox1.Name = "groupBox1";
            groupBox1.Size = new Size(443, 282);
            groupBox1.TabIndex = 10;
            groupBox1.TabStop = false;
            groupBox1.Text = "Loom List";
            // 
            // chkCatLoomList
            // 
            chkCatLoomList.CheckOnClick = true;
            chkCatLoomList.Font = new Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point, 0);
            chkCatLoomList.FormattingEnabled = true;
            chkCatLoomList.Location = new Point(6, 73);
            chkCatLoomList.Name = "chkCatLoomList";
            chkCatLoomList.ScrollAlwaysVisible = true;
            chkCatLoomList.Size = new Size(420, 200);
            chkCatLoomList.TabIndex = 5;
            // 
            // CATLoomWindow
            // 
            AutoScaleDimensions = new SizeF(10F, 25F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(948, 521);
            Controls.Add(groupBox1);
            Controls.Add(Browser);
            Controls.Add(pictureBox1);
            Controls.Add(catFilelbl);
            Controls.Add(catLoomBtn);
            Controls.Add(catLoomTxt);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            Icon = (Icon)resources.GetObject("$this.Icon");
            MaximizeBox = false;
            Name = "CATLoomWindow";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "CAT Loom Wise";
            Load += CATLoomWindow_Load;
            ((System.ComponentModel.ISupportInitialize)pictureBox1).EndInit();
            groupBox1.ResumeLayout(false);
            groupBox1.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private TextBox catLoomTxt;
        private Button catLoomBtn;
        private TextBox catLoomSearchTxt;
        private CheckBox catSearchChkBox;
        private Label catFilelbl;
        private PictureBox pictureBox1;
        private Label Browser;
        private GroupBox groupBox1;
        private CheckedListBox chkCatLoomList;
    }
}
