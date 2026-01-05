namespace Electre_Customize_DotNet.Forms
{
    partial class ExportToPDF
    {

        public TextBox PdfLocationTextBox
        {
            get { return pdflocationText; }
            set { pdflocationText = value; }
        }

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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ExportToPDF));
            label1 = new Label();
            pdflocationText = new TextBox();
            msLoomPicture = new PictureBox();
            generatebtn = new Button();
            ((System.ComponentModel.ISupportInitialize)msLoomPicture).BeginInit();
            SuspendLayout();
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            label1.ForeColor = Color.Blue;
            label1.Location = new Point(19, 41);
            label1.Name = "label1";
            label1.Size = new Size(278, 28);
            label1.TabIndex = 9;
            label1.Text = "BundleReports File Location";
            // 
            // pdflocationText
            // 
            pdflocationText.AccessibleRole = AccessibleRole.Text;
            pdflocationText.Enabled = false;
            pdflocationText.Location = new Point(19, 94);
            pdflocationText.Multiline = true;
            pdflocationText.Name = "pdflocationText";
            pdflocationText.ReadOnly = true;
            pdflocationText.Size = new Size(642, 70);
            pdflocationText.TabIndex = 10;
            pdflocationText.TabStop = false;
            // 
            // msLoomPicture
            // 
            msLoomPicture.Image = (Image)resources.GetObject("msLoomPicture.Image");
            msLoomPicture.Location = new Point(713, 89);
            msLoomPicture.Name = "msLoomPicture";
            msLoomPicture.Size = new Size(150, 75);
            msLoomPicture.SizeMode = PictureBoxSizeMode.StretchImage;
            msLoomPicture.TabIndex = 11;
            msLoomPicture.TabStop = false;
            msLoomPicture.Click += BrowseFile_Click;
            // 
            // generatebtn
            // 
            generatebtn.Font = new Font("Segoe UI", 16F, FontStyle.Bold, GraphicsUnit.Point, 0);
            generatebtn.ForeColor = Color.Blue;
            generatebtn.Location = new Point(249, 261);
            generatebtn.Name = "generatebtn";
            generatebtn.Size = new Size(388, 64);
            generatebtn.TabIndex = 12;
            generatebtn.Text = "Generate PDF";
            generatebtn.UseVisualStyleBackColor = true;
            generatebtn.Click += generatebtnClick;
            // 
            // ExportToPDF
            // 
            AutoScaleDimensions = new SizeF(10F, 25F);
            AutoScaleMode = AutoScaleMode.Font;
            Controls.Add(generatebtn);
            Controls.Add(msLoomPicture);
            Controls.Add(pdflocationText);
            Controls.Add(label1);
            Name = "ExportToPDF";
            Size = new Size(889, 370);
            Load += ExportToPDF_Load;
            ((System.ComponentModel.ISupportInitialize)msLoomPicture).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Label label1;
        private TextBox pdflocationText;
        private PictureBox msLoomPicture;
        private Button generatebtn;
    }
}
