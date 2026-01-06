namespace Electre_Customize_DotNet.ReportUI
{
    partial class PanelDrawingSchedulesWindow
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
            chkSheet = new CheckBox();
            chkListSheets = new CheckedListBox();
            btnFilter = new Button();
            label1 = new Label();
            listBox1 = new ListBox();
            listBox2 = new ListBox();
            label2 = new Label();
            labelMegger = new Label();
            label4 = new Label();
            labelPowerOn = new Label();
            lblLoomCL = new Label();
            lblLoomMegger = new Label();
            lblLoomCont = new Label();
            lblSheetCL = new Label();
            lblPanel = new Label();
            chklistLoomCL = new CheckedListBox();
            chkListSheetCL = new CheckedListBox();
            chkListLoomCont = new CheckedListBox();
            chkListLoomMeg = new CheckedListBox();
            chkListSheetPowerOn = new CheckedListBox();
            chkLoomCL = new CheckBox();
            chkSheetCL = new CheckBox();
            chkSheetPower = new CheckBox();
            chkLoomMeg = new CheckBox();
            chkLoomCon = new CheckBox();
            groupBox1.SuspendLayout();
            SuspendLayout();
            // 
            // groupBox1
            // 
            groupBox1.Controls.Add(chkSheet);
            groupBox1.Controls.Add(chkListSheets);
            groupBox1.Location = new Point(181, 12);
            groupBox1.Name = "groupBox1";
            groupBox1.Size = new Size(443, 160);
            groupBox1.TabIndex = 5;
            groupBox1.TabStop = false;
            groupBox1.Text = "Sheet List";
            // 
            // chkSheet
            // 
            chkSheet.AutoSize = true;
            chkSheet.Cursor = Cursors.Hand;
            chkSheet.Font = new Font("Segoe UI", 9.75F, FontStyle.Bold);
            chkSheet.Location = new Point(291, -4);
            chkSheet.Margin = new Padding(4, 5, 4, 5);
            chkSheet.Name = "chkSheet";
            chkSheet.Size = new Size(127, 32);
            chkSheet.TabIndex = 51;
            chkSheet.Text = "Select All";
            chkSheet.UseVisualStyleBackColor = true;
            chkSheet.CheckedChanged += chkSheet_CheckedChanged;
            // 
            // chkListSheets
            // 
            chkListSheets.CheckOnClick = true;
            chkListSheets.FormattingEnabled = true;
            chkListSheets.Location = new Point(12, 28);
            chkListSheets.Name = "chkListSheets";
            chkListSheets.ScrollAlwaysVisible = true;
            chkListSheets.Size = new Size(391, 116);
            chkListSheets.TabIndex = 1;
            chkListSheets.ItemCheck += chkListSheet_ItemCheck;
            // 
            // btnFilter
            // 
            btnFilter.BackColor = SystemColors.ActiveCaption;
            btnFilter.Font = new Font("Segoe UI", 10F, FontStyle.Bold, GraphicsUnit.Point, 0);
            btnFilter.Location = new Point(630, 122);
            btnFilter.Name = "btnFilter";
            btnFilter.Size = new Size(115, 50);
            btnFilter.TabIndex = 0;
            btnFilter.Text = "Filter";
            btnFilter.UseVisualStyleBackColor = false;
            btnFilter.Click += btnFilter_Click;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(68, 210);
            label1.Name = "label1";
            label1.Size = new Size(59, 25);
            label1.TabIndex = 6;
            label1.Text = "label1";
            // 
            // listBox1
            // 
            listBox1.FormattingEnabled = true;
            listBox1.ItemHeight = 25;
            listBox1.Location = new Point(23, 180);
            listBox1.Name = "listBox1";
            listBox1.Size = new Size(415, 454);
            listBox1.TabIndex = 7;
            // 
            // listBox2
            // 
            listBox2.FormattingEnabled = true;
            listBox2.ItemHeight = 25;
            listBox2.Location = new Point(460, 180);
            listBox2.Name = "listBox2";
            listBox2.Size = new Size(409, 454);
            listBox2.TabIndex = 8;
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Font = new Font("Segoe UI", 10F, FontStyle.Bold, GraphicsUnit.Point, 0);
            label2.Location = new Point(48, 192);
            label2.Name = "label2";
            label2.Size = new Size(103, 28);
            label2.TabIndex = 9;
            label2.Text = "Cable List";
            // 
            // labelMegger
            // 
            labelMegger.AutoSize = true;
            labelMegger.Font = new Font("Segoe UI", 10F, FontStyle.Bold, GraphicsUnit.Point, 0);
            labelMegger.Location = new Point(488, 418);
            labelMegger.Name = "labelMegger";
            labelMegger.Size = new Size(85, 28);
            labelMegger.TabIndex = 17;
            labelMegger.Text = "Megger";
            // 
            // label4
            // 
            label4.AutoSize = true;
            label4.Font = new Font("Segoe UI", 10F, FontStyle.Bold, GraphicsUnit.Point, 0);
            label4.Location = new Point(488, 197);
            label4.Name = "label4";
            label4.Size = new Size(111, 28);
            label4.TabIndex = 26;
            label4.Text = "Continuity";
            // 
            // labelPowerOn
            // 
            labelPowerOn.AutoSize = true;
            labelPowerOn.Font = new Font("Segoe UI", 10F, FontStyle.Bold, GraphicsUnit.Point, 0);
            labelPowerOn.Location = new Point(48, 483);
            labelPowerOn.Name = "labelPowerOn";
            labelPowerOn.Size = new Size(97, 28);
            labelPowerOn.TabIndex = 33;
            labelPowerOn.Text = "PowerOn";
            // 
            // lblLoomCL
            // 
            lblLoomCL.AutoSize = true;
            lblLoomCL.Location = new Point(48, 269);
            lblLoomCL.Name = "lblLoomCL";
            lblLoomCL.Size = new Size(58, 25);
            lblLoomCL.TabIndex = 36;
            lblLoomCL.Text = "Loom";
            // 
            // lblLoomMegger
            // 
            lblLoomMegger.AutoSize = true;
            lblLoomMegger.Location = new Point(488, 542);
            lblLoomMegger.Name = "lblLoomMegger";
            lblLoomMegger.Size = new Size(58, 25);
            lblLoomMegger.TabIndex = 37;
            lblLoomMegger.Text = "Loom";
            // 
            // lblLoomCont
            // 
            lblLoomCont.AutoSize = true;
            lblLoomCont.Location = new Point(488, 286);
            lblLoomCont.Name = "lblLoomCont";
            lblLoomCont.Size = new Size(58, 25);
            lblLoomCont.TabIndex = 38;
            lblLoomCont.Text = "Loom";
            // 
            // lblSheetCL
            // 
            lblSheetCL.AutoSize = true;
            lblSheetCL.Location = new Point(48, 408);
            lblSheetCL.Name = "lblSheetCL";
            lblSheetCL.Size = new Size(56, 25);
            lblSheetCL.TabIndex = 39;
            lblSheetCL.Text = "Sheet";
            // 
            // lblPanel
            // 
            lblPanel.AutoSize = true;
            lblPanel.Location = new Point(48, 550);
            lblPanel.Name = "lblPanel";
            lblPanel.Size = new Size(56, 25);
            lblPanel.TabIndex = 40;
            lblPanel.Text = "Sheet";
            // 
            // chklistLoomCL
            // 
            chklistLoomCL.CheckOnClick = true;
            chklistLoomCL.Font = new Font("Segoe UI", 10F);
            chklistLoomCL.FormattingEnabled = true;
            chklistLoomCL.Location = new Point(138, 239);
            chklistLoomCL.Margin = new Padding(4, 5, 4, 5);
            chklistLoomCL.Name = "chklistLoomCL";
            chklistLoomCL.ScrollAlwaysVisible = true;
            chklistLoomCL.Size = new Size(286, 97);
            chklistLoomCL.TabIndex = 41;
            chklistLoomCL.ItemCheck += chkListLoomCL_ItemCheck;
            // 
            // chkListSheetCL
            // 
            chkListSheetCL.CheckOnClick = true;
            chkListSheetCL.Font = new Font("Segoe UI", 10F);
            chkListSheetCL.FormattingEnabled = true;
            chkListSheetCL.Location = new Point(138, 379);
            chkListSheetCL.Margin = new Padding(4, 5, 4, 5);
            chkListSheetCL.Name = "chkListSheetCL";
            chkListSheetCL.ScrollAlwaysVisible = true;
            chkListSheetCL.Size = new Size(286, 97);
            chkListSheetCL.TabIndex = 42;
            chkListSheetCL.ItemCheck += chkListSheetCL_ItemCheck;
            // 
            // chkListLoomCont
            // 
            chkListLoomCont.CheckOnClick = true;
            chkListLoomCont.Font = new Font("Segoe UI", 10F);
            chkListLoomCont.FormattingEnabled = true;
            chkListLoomCont.Location = new Point(569, 254);
            chkListLoomCont.Margin = new Padding(4, 5, 4, 5);
            chkListLoomCont.Name = "chkListLoomCont";
            chkListLoomCont.ScrollAlwaysVisible = true;
            chkListLoomCont.Size = new Size(289, 128);
            chkListLoomCont.TabIndex = 43;
            chkListLoomCont.ItemCheck += chkListLoomCon_ItemCheck;
            // 
            // chkListLoomMeg
            // 
            chkListLoomMeg.CheckOnClick = true;
            chkListLoomMeg.Font = new Font("Segoe UI", 10F);
            chkListLoomMeg.FormattingEnabled = true;
            chkListLoomMeg.Location = new Point(569, 495);
            chkListLoomMeg.Margin = new Padding(4, 5, 4, 5);
            chkListLoomMeg.Name = "chkListLoomMeg";
            chkListLoomMeg.ScrollAlwaysVisible = true;
            chkListLoomMeg.Size = new Size(289, 128);
            chkListLoomMeg.TabIndex = 44;
            chkListLoomMeg.ItemCheck += chkListLoomMeg_ItemCheck;
            // 
            // chkListSheetPowerOn
            // 
            chkListSheetPowerOn.CheckOnClick = true;
            chkListSheetPowerOn.Font = new Font("Segoe UI", 10F);
            chkListSheetPowerOn.FormattingEnabled = true;
            chkListSheetPowerOn.Location = new Point(138, 530);
            chkListSheetPowerOn.Margin = new Padding(4, 5, 4, 5);
            chkListSheetPowerOn.Name = "chkListSheetPowerOn";
            chkListSheetPowerOn.ScrollAlwaysVisible = true;
            chkListSheetPowerOn.Size = new Size(286, 97);
            chkListSheetPowerOn.TabIndex = 45;
            chkListSheetPowerOn.ItemCheck += chkListSheetPowerOn_ItemCheck;
            // 
            // chkLoomCL
            // 
            chkLoomCL.AutoSize = true;
            chkLoomCL.Cursor = Cursors.Hand;
            chkLoomCL.Font = new Font("Segoe UI", 9.75F, FontStyle.Bold);
            chkLoomCL.Location = new Point(292, 205);
            chkLoomCL.Margin = new Padding(4, 5, 4, 5);
            chkLoomCL.Name = "chkLoomCL";
            chkLoomCL.Size = new Size(127, 32);
            chkLoomCL.TabIndex = 46;
            chkLoomCL.Text = "Select All";
            chkLoomCL.UseVisualStyleBackColor = true;
            chkLoomCL.CheckedChanged += chkLoomCL_CheckedChanged;
            // 
            // chkSheetCL
            // 
            chkSheetCL.AutoSize = true;
            chkSheetCL.Cursor = Cursors.Hand;
            chkSheetCL.Font = new Font("Segoe UI", 9.75F, FontStyle.Bold);
            chkSheetCL.Location = new Point(292, 344);
            chkSheetCL.Margin = new Padding(4, 5, 4, 5);
            chkSheetCL.Name = "chkSheetCL";
            chkSheetCL.Size = new Size(127, 32);
            chkSheetCL.TabIndex = 47;
            chkSheetCL.Text = "Select All";
            chkSheetCL.UseVisualStyleBackColor = true;
            chkSheetCL.CheckedChanged += chkSheetCL_CheckedChanged;
            // 
            // chkSheetPower
            // 
            chkSheetPower.AutoSize = true;
            chkSheetPower.Cursor = Cursors.Hand;
            chkSheetPower.Font = new Font("Segoe UI", 9.75F, FontStyle.Bold);
            chkSheetPower.Location = new Point(292, 492);
            chkSheetPower.Margin = new Padding(4, 5, 4, 5);
            chkSheetPower.Name = "chkSheetPower";
            chkSheetPower.Size = new Size(127, 32);
            chkSheetPower.TabIndex = 48;
            chkSheetPower.Text = "Select All";
            chkSheetPower.UseVisualStyleBackColor = true;
            chkSheetPower.CheckedChanged += chkSheetPowerOn_CheckedChanged;
            // 
            // chkLoomMeg
            // 
            chkLoomMeg.AutoSize = true;
            chkLoomMeg.Cursor = Cursors.Hand;
            chkLoomMeg.Font = new Font("Segoe UI", 9.75F, FontStyle.Bold);
            chkLoomMeg.Location = new Point(731, 456);
            chkLoomMeg.Margin = new Padding(4, 5, 4, 5);
            chkLoomMeg.Name = "chkLoomMeg";
            chkLoomMeg.Size = new Size(127, 32);
            chkLoomMeg.TabIndex = 49;
            chkLoomMeg.Text = "Select All";
            chkLoomMeg.UseVisualStyleBackColor = true;
            chkLoomMeg.CheckedChanged += chkLoomMeg_CheckedChanged;
            // 
            // chkLoomCon
            // 
            chkLoomCon.AutoSize = true;
            chkLoomCon.Cursor = Cursors.Hand;
            chkLoomCon.Font = new Font("Segoe UI", 9.75F, FontStyle.Bold);
            chkLoomCon.Location = new Point(731, 216);
            chkLoomCon.Margin = new Padding(4, 5, 4, 5);
            chkLoomCon.Name = "chkLoomCon";
            chkLoomCon.Size = new Size(127, 32);
            chkLoomCon.TabIndex = 50;
            chkLoomCon.Text = "Select All";
            chkLoomCon.UseVisualStyleBackColor = true;
            chkLoomCon.CheckedChanged += chkLoomCon_CheckedChanged;
            // 
            // PanelDrawingSchedulesWindow
            // 
            AutoScaleDimensions = new SizeF(10F, 25F);
            AutoScaleMode = AutoScaleMode.Font;
            Controls.Add(btnFilter);
            Controls.Add(chkLoomCon);
            Controls.Add(chkLoomMeg);
            Controls.Add(chkSheetPower);
            Controls.Add(chkSheetCL);
            Controls.Add(chkLoomCL);
            Controls.Add(chkListSheetPowerOn);
            Controls.Add(chkListLoomMeg);
            Controls.Add(chkListLoomCont);
            Controls.Add(chkListSheetCL);
            Controls.Add(chklistLoomCL);
            Controls.Add(lblPanel);
            Controls.Add(lblSheetCL);
            Controls.Add(lblLoomCont);
            Controls.Add(lblLoomMegger);
            Controls.Add(lblLoomCL);
            Controls.Add(labelPowerOn);
            Controls.Add(label4);
            Controls.Add(labelMegger);
            Controls.Add(label2);
            Controls.Add(listBox2);
            Controls.Add(listBox1);
            Controls.Add(label1);
            Controls.Add(groupBox1);
            Name = "PanelDrawingSchedulesWindow";
            Size = new Size(893, 640);
            groupBox1.ResumeLayout(false);
            groupBox1.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion
        private GroupBox groupBox1;
        private Button btnFilter;
        private Label label1;
        private ListBox listBox1;
        private ListBox listBox2;
        private Label label2;
        private Label labelMegger;
        private Label label4;
        private CheckedListBox chkListSheets;
        private Label labelPowerOn;
        private Label lblLoomCL;
        private Label lblLoomMegger;
        private Label lblLoomCont;
        private Label lblSheetCL;
        private Label lblPanel;
        public CheckedListBox chklistLoomCL;
        public CheckedListBox chkListSheetCL;
        public CheckedListBox chkListLoomCont;
        private CheckedListBox chkListLoomMeg;
        public CheckedListBox chkListSheetPowerOn;
        public CheckBox chkLoomCL;
        public CheckBox chkSheetCL;
        public CheckBox chkSheetPower;
        public CheckBox chkLoomMeg;
        public CheckBox chkLoomCon;
        public CheckBox chkSheet;
    }
}
