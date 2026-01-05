namespace SegregateLoom
{
    partial class SegregateLoom
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SegregateLoom));
            segraLocationText = new TextBox();
            segregationBtn = new Button();
            label1 = new Label();
            msLoomPicture = new PictureBox();
            label2 = new Label();
            groupBox1 = new GroupBox();
            chckMSLoomListBox = new CheckedListBox();
            msLoomSearchTxt = new TextBox();
            msLoomSegrChkBox = new CheckBox();
            ((System.ComponentModel.ISupportInitialize)msLoomPicture).BeginInit();
            groupBox1.SuspendLayout();
            SuspendLayout();
            // 
            // segraLocationText
            // 
            segraLocationText.Location = new Point(40, 54);
            segraLocationText.Multiline = true;
            segraLocationText.Name = "segraLocationText";
            segraLocationText.Size = new Size(692, 70);
            segraLocationText.TabIndex = 0;
            segraLocationText.TextChanged += segraLocationText_TextChanged;
            // 
            // segregationBtn
            // 
            segregationBtn.AutoSize = true;
            segregationBtn.Cursor = Cursors.Hand;
            segregationBtn.Font = new Font("Segoe UI", 16F, FontStyle.Bold, GraphicsUnit.Point, 0);
            segregationBtn.Location = new Point(235, 451);
            segregationBtn.Name = "segregationBtn";
            segregationBtn.Size = new Size(452, 62);
            segregationBtn.TabIndex = 4;
            segregationBtn.Text = "Generate Loom Wise Report";
            segregationBtn.UseVisualStyleBackColor = true;
            segregationBtn.Click += segregationBtn_Click;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            label1.Location = new Point(40, 25);
            label1.Name = "label1";
            label1.Size = new Size(168, 28);
            label1.TabIndex = 8;
            label1.Text = "MS File Location";
            label1.Click += label1_Click;
            // 
            // msLoomPicture
            // 
            msLoomPicture.Image = (Image)resources.GetObject("msLoomPicture.Image");
            msLoomPicture.Location = new Point(743, 48);
            msLoomPicture.Name = "msLoomPicture";
            msLoomPicture.Size = new Size(150, 75);
            msLoomPicture.SizeMode = PictureBoxSizeMode.StretchImage;
            msLoomPicture.TabIndex = 9;
            msLoomPicture.TabStop = false;
            msLoomPicture.Click += msLoomPicture_Click;
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            label2.Location = new Point(771, 17);
            label2.Name = "label2";
            label2.Size = new Size(89, 28);
            label2.TabIndex = 10;
            label2.Text = "Browser";
            // 
            // groupBox1
            // 
            groupBox1.Controls.Add(chckMSLoomListBox);
            groupBox1.Controls.Add(msLoomSearchTxt);
            groupBox1.Controls.Add(msLoomSegrChkBox);
            groupBox1.Font = new Font("Segoe UI", 12F, FontStyle.Bold, GraphicsUnit.Point, 0);
            groupBox1.ForeColor = Color.Blue;
            groupBox1.Location = new Point(235, 153);
            groupBox1.Name = "groupBox1";
            groupBox1.Size = new Size(452, 292);
            groupBox1.TabIndex = 11;
            groupBox1.TabStop = false;
            groupBox1.Text = "Loom List";
            // 
            // chckMSLoomListBox
            // 
            chckMSLoomListBox.CheckOnClick = true;
            chckMSLoomListBox.Font = new Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point, 0);
            chckMSLoomListBox.FormattingEnabled = true;
            chckMSLoomListBox.Location = new Point(6, 75);
            chckMSLoomListBox.Name = "chckMSLoomListBox";
            chckMSLoomListBox.ScrollAlwaysVisible = true;
            chckMSLoomListBox.Size = new Size(420, 200);
            chckMSLoomListBox.TabIndex = 5;
            // 
            // msLoomSearchTxt
            // 
            msLoomSearchTxt.Font = new Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point, 0);
            msLoomSearchTxt.Location = new Point(6, 38);
            msLoomSearchTxt.Name = "msLoomSearchTxt";
            msLoomSearchTxt.Size = new Size(277, 31);
            msLoomSearchTxt.TabIndex = 2;
            msLoomSearchTxt.TextChanged += msLoomSearchTxt_TextChanged;
            // 
            // msLoomSegrChkBox
            // 
            msLoomSegrChkBox.AutoSize = true;
            msLoomSegrChkBox.Cursor = Cursors.Hand;
            msLoomSegrChkBox.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, 0);
            msLoomSegrChkBox.ForeColor = Color.Blue;
            msLoomSegrChkBox.Location = new Point(309, 38);
            msLoomSegrChkBox.Name = "msLoomSegrChkBox";
            msLoomSegrChkBox.Size = new Size(117, 29);
            msLoomSegrChkBox.TabIndex = 4;
            msLoomSegrChkBox.Text = "Select All";
            msLoomSegrChkBox.UseVisualStyleBackColor = true;
            msLoomSegrChkBox.CheckedChanged += msLoomSegrChkBox_CheckedChanged;
            // 
            // SegregateLoom
            // 
            AutoScaleDimensions = new SizeF(10F, 25F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(905, 525);
            Controls.Add(groupBox1);
            Controls.Add(label2);
            Controls.Add(msLoomPicture);
            Controls.Add(label1);
            Controls.Add(segregationBtn);
            Controls.Add(segraLocationText);
            ForeColor = Color.Blue;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            Icon = (Icon)resources.GetObject("$this.Icon");
            MaximizeBox = false;
            Name = "SegregateLoom";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "MS Loom Wise";
            Load += SegregateLoom_Load;
            ((System.ComponentModel.ISupportInitialize)msLoomPicture).EndInit();
            groupBox1.ResumeLayout(false);
            groupBox1.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }



        #endregion

        private TextBox segraLocationText;
        private Button segregationBtn;
        private Label label1;
        private PictureBox msLoomPicture;
        private Label label2;
        private GroupBox groupBox1;
        private CheckedListBox chckMSLoomListBox;
        private TextBox msLoomSearchTxt;
        private CheckBox msLoomSegrChkBox;
    }
}
