using System;
using System.Data;
using System.Diagnostics;
using System.Linq.Expressions;
using System.Windows.Forms;
using static System.ComponentModel.Design.ObjectSelectorEditor;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace CATLoom
{    
    public partial class CATLoomWindow : Form
    {
        private List<string?> loadedReport = new List<string?>();
        private DataTable dataTable = new DataTable();
        public bool value = false;
        private List<string> checkedItemsAcrossSearches = new List<string>();
        private List<string> checkedItems = new List<string>();
        public CATLoomWindow()
        {
            InitializeComponent();
        }


        // Starts when click on "Generate LoomWise report", the report will start to generate with the selectedItems in the UI
        private void catLoomBtn_Click(object sender, EventArgs e)
        {
            var selectredItems = new List<string>();
            for (int i = 0; i < chkCatLoomList.Items.Count; i++) 
            { 
                if(chkCatLoomList.GetItemChecked(i))
                {
                    var item = chkCatLoomList.Items[i];
                    if(item != null)
                    {
                        selectredItems.Add(item.ToString() ?? string.Empty);
                    }
                }
            }
            generateCATReports(dataTable, selectredItems);
        }

        // create the CATLoom folder under templ and return the destinationFolder.
        private string getDestinationFolder()
        {
            string destinationFolder;
            string tempFloder;
            string fileLocation = Path.GetDirectoryName(catLoomTxt.Text);
            string projectFolder = Path.GetDirectoryName(fileLocation).ToLower();
            try
            {

                if (GlobalVar.StrtCmd.ToLower() == projectFolder)
                {
                    tempFloder = Path.Combine(GlobalVar.StrtCmd, "templ");
                    destinationFolder = Path.Combine(tempFloder, "CATLOOM");
                }

                else
                {
                    tempFloder = Path.Combine(fileLocation, "templ");
                    destinationFolder = Path.Combine(tempFloder, "CATLOOM");
                }


                if (!Directory.Exists(tempFloder))
                {
                    Directory.CreateDirectory(tempFloder);
                }

                if (!Directory.Exists(destinationFolder))
                {
                    Directory.CreateDirectory(destinationFolder);
                }
                return destinationFolder;
            }
            catch (Exception ex)
            {

                Logging.Error(ex.Message);
            }
            return "";
        }

        // Reads a CSV file (specified by the filePath), parses the data, and returns the data in the form of a DataTable
        private DataTable ReadCsv(string filePath)
        {
            DataTable dataTablefunc = new DataTable();
            try
            {
                using (StreamReader reader = new StreamReader(filePath))
                {
                    string[] headers = reader.ReadLine().Split(';');
                    foreach (string header in headers)
                    {
                        dataTablefunc.Columns.Add(header);
                    }

                    while (!reader.EndOfStream)
                    {
                        string[] rows = reader.ReadLine().Split(';');
                        dataTablefunc.Rows.Add(rows);
                    }
                }

            }
            catch (Exception ex)
            {
                Logging.Error(ex.Message);

            }
            return dataTablefunc;

        }
        // Takes the data stored in a DataTable (which contains rows and columns) and writes it to a CSV file at the specified filePath
        private void WriteCsv(DataTable dataTable, string filePath)
        {
            try
            {

                if (Directory.Exists(filePath))
                {
                    Directory.Delete(filePath, true);
                }



                using (StreamWriter writer = new StreamWriter(filePath))
                {
                    // Write the header row
                    writer.WriteLine(string.Join(";", dataTable.Columns.Cast<DataColumn>().Select(col => col.ColumnName)));

                    // Write data rows
                    foreach (DataRow row in dataTable.Rows)
                    {
                        writer.WriteLine(string.Join(";", row.ItemArray));
                    }
                }
            }
            catch (Exception ex)
            {
                Logging.Error(ex.Message);
            }
        }
        // Read data from a CSV file, search for Bundle_from column exists, Extract unique values from the "Bundle_from" column of the CSV file and Populate the ListBox with those unique values.
        private void loadSearchCATList()
        {
            try
            {
                chkCatLoomList.Items.Clear();
                string filePath = catLoomTxt.Text;

                if (!string.IsNullOrEmpty(filePath))
                {
                    dataTable = ReadCsv(filePath);
                    

                    if (!dataTable.Columns.Contains("Bundle_from"))
                    {
                        MessageBox.Show("The column 'Bundle_from' does not exist in the selected CSV file.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    List<string?> uniqueValues = dataTable.AsEnumerable()
                                                .Where(row => !row.IsNull("Bundle_from"))
                                                .Select(row => row.Field<string>("Bundle_from"))
                                                .Distinct()
                                                .OrderBy(val => val)
                                                .ToList();

                    loadedReport.Clear();
                    loadedReport.AddRange(uniqueValues);
                    populateCatList(loadedReport);
                }

            }

            catch (IOException ex)
            {
                MessageBox.Show("File Open in another Application\n Please Close and Try again", "Report Generating Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Logging.Error(ex.Message);
            }

            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Logging.Error(ex.Message);
            }
        }
        // The method generates and saves CSV reports for selected items(looms) from the data provided in the DataTable.
// It filters the data based on the "Bundle_from" column and writes each filtered set of data to a separate CSV file.
        private void generateCATReports(DataTable dataTable, List<string> selectedReport)
        {
            //   string directortyPath = Path.GetDirectoryName(filePath);

            try
            {

                string outputFolder = getDestinationFolder();
                Logging.Info($"Output folder selected: {outputFolder}");

                if (selectedReport.Count != 0)
                {
                    Logging.Info($"Total selected reports: {selectedReport.Count}");

                    // Group dataTable's rows by "Bundle_from" once instead of rescanning the whole
                    // table for every selected loom (was O(selectedReport.Count * dataTable.Rows.Count)).
                    // ToLookup preserves each group's original row order, same as the per-value
                    // .Where() did, and a missing key still yields an empty sequence (so
                    // CopyToDataTable() on a loom with no matching rows still throws into the same
                    // per-value catch below, exactly as before).
                    var rowsByBundleFrom = dataTable.AsEnumerable().ToLookup(row => row.Field<string>("Bundle_from"));

                    foreach (string value in selectedReport)
                    {
                        try
                        {
                            Logging.Info($"Processing loom: {value}");
                            var filteredData = rowsByBundleFrom[value].CopyToDataTable();

                            string outputFileName = Path.Combine(outputFolder, $"{value}.csv");
                            WriteCsv(filteredData, outputFileName);
                            Logging.Info($"Successfully created report for loom: {value}. File: {outputFileName}");
                        }
                        catch (System.Exception ex)
                        {
                            Logging.Error($"Error processing loom: {value}. Error: {ex.Message}");

                        }
                    }
                    var result = MessageBox.Show($"CAT_LOOM files created successfully in the folder: {outputFolder}\nClick Ok to view the folder path", "Success", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                    if (result == DialogResult.OK)
                    {
                        System.Diagnostics.Process.Start("explorer.exe", outputFolder);
                    }

                    Logging.Error("All selected reports generated successfully.");
                    this.Close();
                }

                else
                {
                    MessageBox.Show("No Data on the Bundle Looms", "Failed", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Logging.Error("No looms selected or data unavailable.");
                    this.Close();
                }
            }
            catch (Exception ex)
            {

                Logging.Error(ex.Message);
            }
        }

        // populates the chkCatLoomList ListBox with the provided list of strings
        private void populateCatList(List<string> catList)
        {
            foreach (string cat in catList)
            {
                chkCatLoomList.Items.Add(cat);
            }
        }
        // search and filter the list of items (looms) while preserving the checked items that the user has selected.
        private void searchCatItem(string searchTerm, List<string> catReportList)
        {
            List<string> checkedItemsBeforeSearch = chkCatLoomList.CheckedItems.Cast<string>().ToList();
            chkCatLoomList.Items.Clear();
            List<string> itemsToAdd;
            if (string.IsNullOrEmpty(searchTerm))
            {
                itemsToAdd = catReportList.ToList();
            }
            else
            {
                itemsToAdd = catReportList.Where(e => e.ToLower().Contains(searchTerm.ToLower())).ToList();
            }

            chkCatLoomList.Items.AddRange(itemsToAdd.ToArray());

            foreach (var checkedItem in checkedItemsBeforeSearch)
            {
                if (!checkedItemsAcrossSearches.Contains(checkedItem))
                {
                    checkedItemsAcrossSearches.Add(checkedItem);
                }
            }

            foreach(var checkedItem in checkedItemsAcrossSearches)
            {
                int index = chkCatLoomList.Items.IndexOf(checkedItem);
                if(index >= 0)
                {
                    chkCatLoomList.SetItemChecked(index, true);
                }
            }

            catSearchChkBox.Checked = false;
        }
        // Loads the cat_data.csv file
        private void CATLoomWindow_Load(object sender, EventArgs e)
        {
            try
            {
                string catLoomDefaulPath = Path.Combine(GlobalVar.StrtCmd, "CAT_data.csv");

                if (!File.Exists(catLoomDefaulPath))
                {
                    MessageBox.Show("CAT_LOOM.csv report not Exist in project.\n", "Report Missing", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    catLoomTxt.Text = "";
                }

                else
                {
                    catLoomTxt.Text = catLoomDefaulPath;
                }


                catLoomTxt.Enabled = false;
                loadSearchCATList();
            }
            catch (Exception ex)
            {

                Logging.Error(ex.Message);
            }

        }
        // Search items (looms)
        private void catLoomSearchTxt_TextChanged(object sender, EventArgs e)
        {
            searchCatItem(catLoomSearchTxt.Text.ToLower().Trim(), loadedReport);
            selectSearchedItems(chkCatLoomList, catSearchChkBox);
        }
        // select/ deselect all items (looms)
        private void catSearchChkBox_CheckedChanged(object sender, EventArgs e)
        {
            //selectSearchedItems(catLoomList, catSearchChkBox);
            if (catSearchChkBox.Checked) 
            { 
                for(int i = 0; i < chkCatLoomList.Items.Count; i++)
                {
                    chkCatLoomList.SetItemChecked(i, true);
                }
                checkedItems = chkCatLoomList.Items.Cast<string>().ToList();
            }
            else
            {
                for(int i = 0; i < chkCatLoomList.Items.Count; i++)
                {
                    chkCatLoomList.SetItemChecked(i, false);
                }

                checkedItemsAcrossSearches.Clear();
                checkedItems.Clear();
            }
        }
        // Selects/Clears all items in the ListBox and enables/disables the ListBox again
        private void selectSearchedItems(ListBox listbox, CheckBox checkbox)
        {

            if (checkbox.Checked == true)
            {
                bool isChecked = checkbox.Checked;
                for (int i = 0; i < listbox.Items.Count; i++)
                {
                    listbox.SetSelected(i, isChecked);
                }

                listbox.Enabled = false;

                value = true;
            }

            else
            {
                listbox.ClearSelected();
                listbox.Enabled = true;
            }
            listbox.Enabled = true;
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
        // Allows the user to select a CSV file from their computer and then loads the selected file into the application

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            try
            {
                using (OpenFileDialog openFileDialog = new OpenFileDialog())
                {

                    openFileDialog.InitialDirectory = GlobalVar.StrtCmd; // Replace with your desired default folder


                    openFileDialog.Filter = "CSV files (*.csv)|*.csv";

                    DialogResult result = openFileDialog.ShowDialog();

                    if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(openFileDialog.FileName))
                    {
                        catLoomTxt.Text = openFileDialog.FileName;
                    }
                }
                loadSearchCATList();
                catSearchChkBox.Checked = false;
                selectSearchedItems(chkCatLoomList, catSearchChkBox);
            }
            catch (Exception ex)
            {
                Logging.Error(ex.Message);

            }
        }
        // not using
        private void catLoomList_DrawItem(object sender, DrawItemEventArgs e)
        {
            selection_color(sender, e);
        }
        // not using
        private void selection_color(object sender, DrawItemEventArgs e)
        {
            ListBox listBox = (ListBox)sender;


            if (e.Index < 0) return;
            string itemText = listBox.Items[e.Index].ToString();

            bool isSelected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;

            Color backgroundColor = isSelected ? Color.LightGreen : Color.White;
            Color textColor = isSelected ? Color.Black : Color.Black;

            using (SolidBrush backgroundBrush = new SolidBrush(backgroundColor))
            {
                e.Graphics.FillRectangle(backgroundBrush, e.Bounds);
            }

            using (SolidBrush textBrush = new SolidBrush(textColor))
            {
                e.Graphics.DrawString(itemText, e.Font, textBrush, e.Bounds);
            }

            e.DrawFocusRectangle();
        }
        private void catLoomTxt_TextChanged(object sender, EventArgs e)
        {

        }
        private void catLoomList_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }     
    }
}
 