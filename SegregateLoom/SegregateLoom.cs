using System.Diagnostics;
using System.DirectoryServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
//using ClosedXML.Excel;

namespace SegregateLoom
{
    public partial class SegregateLoom : Form
    {

        List<string> matList = new List<string>();
        private List<string> checkedItemsAcrossSearches = new List<string>();
        private List<string> checkedItems = new List<string>();
        public bool value = false;

        public SegregateLoom()
        {
            InitializeComponent();
        }

        private void locationBtn_Click(object sender, EventArgs e)
        {

        }
        // Starts when click on "Generate Loom wise report", the report will start to generate with the selectedItems in the UI
        private void segregationBtn_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            var selectedItems = new List<string>();
            for(int i = 0; i <chckMSLoomListBox.Items.Count; i++)
            {
                if (chckMSLoomListBox.GetItemChecked(i))
                {
                    var item = chckMSLoomListBox.Items[i];
                    if(item != null)
                    {
                        selectedItems.Add(item.ToString() ?? string.Empty);
                    }
                }
            }
            generateReport(selectedItems);
            Cursor = Cursors.Default;
            Application.Exit();
        }

        // create the MS loom folder under templ and return the destinationFolder.
        private string getDestinationFolder()
        {
            try
            {
                string destinationFolder;
                string tempFloder;
                string fileLocation = Path.GetDirectoryName(segraLocationText.Text);
                string projectFolder = Path.GetDirectoryName(fileLocation).ToLower();

                string strtCmd = GlobalVar.StrtCmd.ToLower();

                if (strtCmd.EndsWith('\\'))
                {
                    strtCmd = strtCmd.TrimEnd('\\');
                }

                if (strtCmd == projectFolder)
                {
                    tempFloder = Path.Combine(GlobalVar.StrtCmd, "templ");
                    destinationFolder = Path.Combine(tempFloder, "MSLOOM");
                }

                else
                {
                    tempFloder = Path.Combine(fileLocation, "templ");
                    destinationFolder = Path.Combine(tempFloder, "MSLOOM");
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
        // Creates the Excel report for the selected ietms (looms) in the provided destinationFolder with the help of ProcessWorksheet method

        private void generateReport(List<string> searchedList)
        {
            string destinationFolder = getDestinationFolder();
            Logging.Info($"Output folder selected: {destinationFolder}");


            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Open(segraLocationText.Text);

            try
            {
                foreach (Excel.Worksheet worksheet in workbook.Worksheets)
                {
                    if (searchedList.Contains(worksheet.Name.ToUpper()))
                    {
                        ProcessWorksheet(worksheet, destinationFolder);
                    }
                }

                workbook.Close(false);
                excelApp.Quit();
            }
            finally
            {
                ReleaseCOM(workbook);
                ReleaseCOM(excelApp);
            }
            try
            {
                var result = MessageBox.Show($"MS_LOOM files created successfully in the folder:\n {destinationFolder}\nClick Ok to view the folder path", "Success", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                if (result == DialogResult.OK)
                {
                    System.Diagnostics.Process.Start("explorer.exe", destinationFolder);
                }
                Logging.Info("MS_LOOM files created successfully");
            }
            catch (IOException ex)
            {
                MessageBox.Show("File Opened in another Application.\nPlease close and Try agin", "Report Generateing Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Logging.Error(ex.Message);
            }

            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}");
                Logging.Error(ex.Message);
                this.Close();
            }
        }

        #region CommentedCode on Dec,05,2025 due to ClosedXML replacement with Interop
        //private void generateReport(List<string> searchedList)
        //{
        //    string destinationFolder = getDestinationFolder();
        //    Logging.Info($"Output folder selected: {destinationFolder}");


        //    using (var workbook = new XLWorkbook(segraLocationText.Text))
        //    {
        //        foreach (var worksheet in workbook.Worksheets)
        //        {
        //            if (searchedList.Exists(e => e == worksheet.Name.ToUpper()))
        //            {
        //                ProcessWorksheet(worksheet, destinationFolder);
        //            }
        //        }
        //    }

        //    try
        //    {                     
        //        var result = MessageBox.Show($"MS_LOOM files created successfully in the folder:\n {destinationFolder}\nClick Ok to view the folder path", "Success", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
        //        if (result == DialogResult.OK)
        //        {
        //            System.Diagnostics.Process.Start("explorer.exe", destinationFolder);
        //        }
        //        Logging.Info("MS_LOOM files created successfully");
        //    }
        //    catch (IOException ex)
        //    {
        //        MessageBox.Show("File Opened in another Application.\nPlease close and Try agin", "Report Generateing Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //        Logging.Error(ex.Message);
        //    }

        //    catch (Exception ex)
        //    {
        //        MessageBox.Show($"An error occurred: {ex.Message}");
        //        Logging.Error(ex.Message);
        //        this.Close();
        //    }
        //}
        // Selects/Clears all items in the ListBox and enables/disables the ListBox again
        #endregion
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
        // loads the mat.xlsx file from the directory
        private void SegregateLoom_Load(object sender, EventArgs e)
        {
            //string msLoomDirectory = Path.GetDirectoryName(GlobalVar.matFolder);

            segraLocationText.Enabled = false;

            //MessageBox.Show("Show", "", MessageBoxButtons.OK); for testing.
            if (Directory.Exists(GlobalVar.MsFolder))
            {
                var MsloomFile = Directory.GetFiles(GlobalVar.MsFolder, "*_mat.xlsx").FirstOrDefault();

                if (!File.Exists(MsloomFile))
                {
                    MessageBox.Show("Material File not exist in the Result Folder.\nPlease generate Metrial List report from OOTB", "File Missing", MessageBoxButtons.OK, MessageBoxIcon.Question);
                    Logging.Info("Material File not exist in the Result Folder");
                    return;
                }
                loadFileCont(MsloomFile);
            }

            else
            {
                MessageBox.Show("Result Folder not exist in the Project. \nKindly generate or select other MSloom file.", "Result Folder Missing", MessageBoxButtons.OK, MessageBoxIcon.Question);
                Logging.Info("Result Folder not exist in the project");
                return;
            }
        }
        // search and filter the list of items (looms) while preserving the checked items that the user has selected.
        private void searchMSItem(string searchTerm, List<string> msReportList)
        {

            List<string> checkedItemsBeforeSearch = chckMSLoomListBox.CheckedItems.Cast<string>().ToList();
            chckMSLoomListBox.Items.Clear();
            List<string> itemsToAdd;
            if (string.IsNullOrEmpty(searchTerm))
            {
                itemsToAdd = msReportList.OrderBy(e => e).ToList();
            }
            else
            {
                itemsToAdd = msReportList.Where(e => e.ToLower().Contains(searchTerm.ToLower())).OrderBy(e => e).ToList();
            }

            chckMSLoomListBox.Items.AddRange(itemsToAdd.ToArray());
            foreach(var checkedItem in checkedItemsBeforeSearch)
            {
                if (!checkedItemsAcrossSearches.Contains(checkedItem))
                {
                    checkedItemsAcrossSearches.Add(checkedItem);
                }
            }

            foreach(var checkedItem in checkedItemsAcrossSearches)
            {
                int index = chckMSLoomListBox.Items.IndexOf(checkedItem);
                if(index >= 0)
                {
                    chckMSLoomListBox.SetItemChecked(index, true);
                }
            }

            msLoomSegrChkBox.Checked = false;
        }

        // search the particular loom through search bar
        private void msLoomSearchTxt_TextChanged(object sender, EventArgs e)
        {
            searchMSItem(msLoomSearchTxt.Text.ToLower().Trim(), matList);
        }
        // select/ deselect all items (looms)
        private void msLoomSegrChkBox_CheckedChanged(object sender, EventArgs e)
        {
            if (msLoomSegrChkBox.Checked)
            {
                for(int i = 0; i < chckMSLoomListBox.Items.Count; i++)
                {
                    chckMSLoomListBox.SetItemChecked(i, true);
                }
                checkedItems = chckMSLoomListBox.Items.Cast<string>().ToList();
            }
            else
            {
                for(int i = 0; i < chckMSLoomListBox.Items.Count; i++)
                {
                    chckMSLoomListBox.SetItemChecked(i, false);
                }

                checkedItemsAcrossSearches.Clear();
                checkedItems.Clear();
            }
        }
        // opens the local folder for listing mat excel files
        private void msLoomPicture_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {

                openFileDialog.InitialDirectory = GlobalVar.MsFolder; // Replace with your desired default folder

                openFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*";

                DialogResult result = openFileDialog.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(openFileDialog.FileName))
                {
                    segraLocationText.Text = openFileDialog.FileName;
                }

                loadFileCont(segraLocationText.Text);
                msLoomSegrChkBox.Checked = false;
                selectSearchedItems(chckMSLoomListBox, msLoomSegrChkBox);
            }
        }
        // loading the data from the selected mat excel file if exists
        private void loadFileCont(string fileLocation)
        {
            chckMSLoomListBox.Items.Clear();
            matList.Clear();
            if (fileLocation != null)
            {
                Excel.Application excelApp = new Excel.Application();

                Excel.Workbook workbook = excelApp.Workbooks.Open(fileLocation);
                foreach (Excel.Worksheet worksheet in workbook.Sheets)
                {
                    matList.Add(worksheet.Name);
                }
                excelApp.Quit();
                var sortedList = matList.OrderBy(name => name).ToList();

                // Add to UI
                chckMSLoomListBox.Items.AddRange(sortedList.ToArray());
                segraLocationText.Text = fileLocation;
            }

            else
            {
                MessageBox.Show("File Not Exitst", "File no Found", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Logging.Error("File Not Exitst");
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
        // not using
        private void msLoomListBox_DrawItem(object sender, DrawItemEventArgs e)
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
        //creating the worksheets for selected items
        private void ProcessWorksheet(Excel.Worksheet worksheet, string outputFolder)
        {
            try
            {
                Excel.Range used = worksheet.UsedRange;

                int lastRowUsed = used.Rows.Count;
                int lastColumnUsed = used.Columns.Count;

                string fileName = Path.Combine(outputFolder, $"{worksheet.Name}_2DMS.csv");

                using (StreamWriter writer = new StreamWriter(fileName))
                {
                    string wsName = worksheet.Name.ToUpper();

                    int startRow = 1;
                    int endRow = lastRowUsed;

                    if (wsName == "MECHANICAL PARTS")
                        startRow = 4;

                    if (wsName != "WEIGHT RESULT" && wsName != "MECHANICAL PARTS")
                        endRow = lastRowUsed - 2;

                    for (int row = startRow; row <= endRow; row++)
                    {
                        string line = "";

                        for (int col = 1; col <= lastColumnUsed; col++)
                        {
                            var cellValue = worksheet.Cells[row, col].Value2;
                            string text = cellValue != null ? cellValue.ToString() : "";

                            line += text + ";";

                            if (text == "Weight" && wsName != "WEIGHT RESULT" && wsName != "MECHANICAL PARTS")
                                line += "Remarks;";
                        }

                        if (wsName == "WEIGHT RESULT" && row == 1)
                            line += "Remarks";

                        if (wsName == "MECHANICAL PARTS" && row == 4)
                            line += "Remarks";

                        writer.WriteLine(line);
                    }
                }

                Logging.Info($"{worksheet.Name}.csv generated");
                ReleaseCOM(used);
            }
            catch (Exception ex)
            {
                Logging.Error(ex.Message);
            }
        }

        #region CommentedCode on Dec,05,2025 due to ClosedXML replacement with Interop
        //private void ProcessWorksheet(IXLWorksheet worksheet, string outputFolder)
        //{
        //    try
        //    {
        //        var lastRowUsed = worksheet.LastRowUsed().RowNumber();
        //        var lastColumnUsed = worksheet.LastColumnUsed().ColumnNumber();

        //        string fileName = Path.Combine(outputFolder, $"{worksheet.Name}_2DMS.csv");
        //        using (StreamWriter writer = new StreamWriter(fileName))
        //        {

        //            if (worksheet.Name.ToUpper() == "WEIGHT RESULT")
        //            {
        //                for (int row = 1; row <= lastRowUsed; row++) // Skip last 2 rows
        //                {
        //                    string line = "";
        //                    for (int col = 1; col <= lastColumnUsed; col++)
        //                    {
        //                        var cellValue = worksheet.Cell(row, col).Value.ToString();
        //                        //   string columnName = worksheet.Cell(1, col).GetString();                    


        //                        line += cellValue + ";";
        //                    }

        //                    // Add Remarks column at the end
        //                    if (row == 1)
        //                    {
        //                        line += "Remarks";
        //                    }

        //                    writer.WriteLine(line);

        //                }
        //                Logging.Info("Weight result.csv file is created");
        //            }

        //            else if (worksheet.Name.ToUpper() == "MECHANICAL PARTS")
        //            {
        //                for (int row = 4; row <= lastRowUsed; row++) // Skip first 4 rows
        //                {
        //                    string line = "";
        //                    for (int col = 1; col <= lastColumnUsed; col++)
        //                    {
        //                        var cellValue = worksheet.Cell(row, col).Value.ToString();
        //                        //   string columnName = worksheet.Cell(1, col).GetString();                    


        //                        line += cellValue + ";";
        //                    }

        //                    // Add Remarks column at the end
        //                    if (row == 4)
        //                    {
        //                        line += "Remarks";
        //                    }

        //                    writer.WriteLine(line);
        //                }
        //                Logging.Info("Mechanical parts.csv file is created");
        //            }

        //            else
        //            {
        //                for (int row = 1; row <= lastRowUsed - 2; row++) // Skip last 2 rows
        //                {
        //                    string line = "";
        //                    for (int col = 1; col <= lastColumnUsed; col++)
        //                    {
        //                        var cellValue = worksheet.Cell(row, col).Value.ToString();
        //                        //   string columnName = worksheet.Cell(1, col).GetString();                    


        //                        line += cellValue + ";";
        //                        // Add Remarks Column between Weight and Effectivity
        //                        if (cellValue == "Weight")
        //                        {
        //                            line += "Remarks"+";";
        //                        }
        //                    }

        //                    // Add Remarks column at the end
        //                    //if (row == 1)
        //                    // {
        //                    //     line += "Remarks";
        //                    // }

        //                    writer.WriteLine(line);

        //                }
        //                Logging.Info($"{worksheet}.csv file created");
        //            }
        //        }
        //    }
        //    catch (Exception e)
        //    {

        //        Logging.Info(e.Message);
        //    }
        //}
        #endregion
        private void ReleaseCOM(object obj)
        {
            try
            {
                if (obj != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
            }
            catch { }
        }

        private void segraLocationText_TextChanged(object sender, EventArgs e)
        {

        }

        private void msLoomListBox_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
        // not using
        private void MsLoomListBox_MouseDown(object sender, MouseEventArgs e)
        {
            ListBox listBox = sender as ListBox;

            if (listBox == null) return;

            // Get the index of the item clicked
            int index = listBox.IndexFromPoint(e.Location);
            if (value == false)
            {


                if (index != ListBox.NoMatches)
                {
                    // Toggle the selection state
                    if (listBox.SelectedIndices.Contains(index))
                    {
                        listBox.SetSelected(index, true);
                    }
                    else
                    {
                        listBox.SetSelected(index, false);
                    }


                }
            }
            else
            {
                listBox.SetSelected(index, false); 
            }
            value = false;

        }
    }
}

