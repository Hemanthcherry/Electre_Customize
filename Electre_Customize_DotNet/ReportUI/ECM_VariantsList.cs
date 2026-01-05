using Electre_Customize_DotNet.Contracts;
using Electre_Customize_DotNet.MainOperation;
using Electre_Customize_DotNet.Objects;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace Electre_Customize_DotNet.ReportUI
{
    public partial class ECM_VariantsList : Form
    {
        //private ISelectionFunctions _selectionFunctions;
        public static string? SelectedVariant { get; private set; }  // To store selected variant name
        public int TotalVariant = 0;
        private List<string?> originalVariantList = new List<string?>();

        public ECM_VariantsList()
        {
            InitializeComponent();
            this.StartPosition = FormStartPosition.CenterScreen;

            //_selectionFunctions = SelectionFunctions.SelectionFuntionsInstance;
            // Load folders
            string projectPath = GlobalVar.StrtCmd;
            if (!string.IsNullOrEmpty(projectPath))
            {
                originalVariantList = Directory.GetDirectories(projectPath, "schem_ecm_copy_*", SearchOption.TopDirectoryOnly)
                                                        .Select(Path.GetFileName)
                                                        .OrderBy(x => x)
                                                        .ToList();
            }
            else
            {
                originalVariantList = new List<string?>();
                MessageBox.Show("Project path is not set. Please check the configuration.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
            }


            chkListVariants.Items.Clear();
            foreach (var variant in originalVariantList)
            {                
                chkListVariants.Items.Add(variant);
            }

            // Update label with total count
            lblTotalVariants.Text = $"Total Variants: {chkListVariants.Items.Count}";

            // Enforce single selection
            chkListVariants.ItemCheck += ChkListVarients_ItemCheck;
        }

        private void ChkListVarients_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            if (e.NewValue == CheckState.Checked)
            {
                // Uncheck all other items
                for (int i = 0; i < chkListVariants.Items.Count; i++)
                {
                    if (i != e.Index)
                    {
                        chkListVariants.SetItemChecked(i, false);
                    }
                }
            }
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            if (chkListVariants.CheckedItems.Count != 1)
            {
                MessageBox.Show("Please select exactly one variant.");
                return;
            }

            // Store the selected variant name
            SelectedVariant = chkListVariants.CheckedItems[0].ToString();

            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void txtVariantFilter_TextChanged(object sender, EventArgs e)
        {
            if (txtVariantFilter.Text.ToUpper() == "SEARCH FILTER") return;

            string filter = txtVariantFilter.Text.ToLower();
            var filteredItems = originalVariantList
                .Where(item => item.ToLower().Contains(filter))
                .ToList();

            string previouslyChecked = null;
            if (chkListVariants.CheckedItems.Count > 0)
            {
                previouslyChecked = chkListVariants.CheckedItems[0].ToString();
            }

            chkListVariants.Items.Clear();
            foreach (var item in filteredItems)
            {
                bool isChecked = (item == previouslyChecked);
                chkListVariants.Items.Add(item, isChecked);
            }

            lblTotalVariants.Text = $"Total Variants: {filteredItems.Count}";
        }

        private void txtVariantFilter_Click(object sender, EventArgs e)
        {
            if (txtVariantFilter.Text == "Search Filter")
            {
                txtVariantFilter.Text = "";
                txtVariantFilter.ForeColor = Color.Black;
                txtVariantFilter.Font = new Font("Segoe UI", 10F);
            }
        }

        private void txtVariantFilter_Leave(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtVariantFilter.Text))
            {
                txtVariantFilter.Text = "Search Filter";
                txtVariantFilter.ForeColor = Color.Silver;
                txtVariantFilter.Font = new Font("Segoe UI", 12F, FontStyle.Italic);
            }
        }
    }
}
