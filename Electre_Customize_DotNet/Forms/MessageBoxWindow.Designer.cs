namespace Electre_Customize_DotNet.Forms
{
    partial class MessageBoxWindow
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MessageBoxWindow));
            wireListBtn = new Button();
            wireList = new Label();
            duplicateWireGridview = new DataGridView();
            ((System.ComponentModel.ISupportInitialize)duplicateWireGridview).BeginInit();
            SuspendLayout();
            // 
            // wireListBtn
            // 
            wireListBtn.BackColor = SystemColors.Control;
            wireListBtn.Font = new Font("Segoe UI", 12F, FontStyle.Bold, GraphicsUnit.Point, 0);
            wireListBtn.ForeColor = Color.Blue;
            wireListBtn.Location = new Point(959, 265);
            wireListBtn.Name = "wireListBtn";
            wireListBtn.Size = new Size(112, 67);
            wireListBtn.TabIndex = 0;
            wireListBtn.Text = "OK";
            wireListBtn.UseVisualStyleBackColor = false;
            wireListBtn.UseWaitCursor = true;
            wireListBtn.Click += wireListBtn_Click;
            // 
            // wireList
            // 
            wireList.AutoSize = true;
            wireList.Font = new Font("Segoe UI Semibold", 12F, FontStyle.Bold, GraphicsUnit.Point, 0);
            wireList.ForeColor = Color.Red;
            wireList.Location = new Point(25, 34);
            wireList.Name = "wireList";
            wireList.Size = new Size(464, 32);
            wireList.TabIndex = 2;
            wireList.Text = "Duplicate Wires are found please resolve.";
            wireList.UseWaitCursor = true;
            // 
            // duplicateWireGridview
            // 
            duplicateWireGridview.AllowUserToAddRows = false;
            duplicateWireGridview.AllowUserToDeleteRows = false;
            duplicateWireGridview.AllowUserToOrderColumns = true;
            duplicateWireGridview.BackgroundColor = Color.White;
            duplicateWireGridview.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;
            duplicateWireGridview.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            duplicateWireGridview.Location = new Point(25, 88);
            duplicateWireGridview.Name = "duplicateWireGridview";
            duplicateWireGridview.RowHeadersWidth = 62;
            duplicateWireGridview.Size = new Size(928, 437);
            duplicateWireGridview.TabIndex = 3;
            duplicateWireGridview.UseWaitCursor = true;
            duplicateWireGridview.CellContentClick += duplicateWireGridview_CellContentClick;
            // 
            // MessageBoxWindow
            // 
            AutoScaleDimensions = new SizeF(10F, 25F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1083, 551);
            Controls.Add(duplicateWireGridview);
            Controls.Add(wireList);
            Controls.Add(wireListBtn);
            FormBorderStyle = FormBorderStyle.FixedSingle;
            Icon = (Icon)resources.GetObject("$this.Icon");
            MaximizeBox = false;
            Name = "MessageBoxWindow";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "Info Message";
            UseWaitCursor = true;
            Load += MessageBoxWindow_Load;
            ((System.ComponentModel.ISupportInitialize)duplicateWireGridview).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button wireListBtn;
        private Label wireList;
        private DataGridView duplicateWireGridview;
    }
}