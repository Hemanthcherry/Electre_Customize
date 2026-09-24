using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;
using Range = Microsoft.Office.Interop.Excel.Range; // For drawing on the PDF

namespace Electre_Customize_DotNet.Forms
{
    public partial class ExportToPDF : UserControl
    {
        public ExportToPDF()
        {
            InitializeComponent();
        }

        private void BrowseFile_Click(object sender, EventArgs e)
        {
            // Create an instance of OpenFileDialog
            OpenFileDialog openFileDialog = new OpenFileDialog();

            // Show the file dialog and check if the user selected a file
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                // If a file was selected, get the file path
                string selectedFilePath = openFileDialog.FileName;

                // Assuming your TextBox is called pdflocationText (or the name of your TextBox)
                pdflocationText.Text = selectedFilePath;  // Set the selected file path in the TextBox
            }

        }

        private void generatebtnClick(object sender, EventArgs e)
        {
            try
            {
                generatebtn.Enabled = false;
                string filePath = pdflocationText.Text;

                // Check if the file has a valid Excel extension (.xls or .xlsx)
                string fileExtension = System.IO.Path.GetExtension(filePath);
                if (!string.Equals(fileExtension, ".xls", StringComparison.OrdinalIgnoreCase)
                    && !string.Equals(fileExtension, ".xlsx", StringComparison.OrdinalIgnoreCase))
                {
                    MessageBox.Show("The selected file is not an Excel file.");
                    generatebtn.Enabled = true;
                    return;
                }

                // Initialize Excel application
                Application excelApp = new Application();
                if (excelApp == null)
                {
                    MessageBox.Show("Excel is not properly installed!");
                    return;
                }

                Workbook workBook = excelApp.Workbooks.Open(filePath);
                Microsoft.Office.Interop.Excel.Worksheet ws1 = (Microsoft.Office.Interop.Excel.Worksheet)workBook.Sheets[1];
                string valueF51 = ws1.Range["F51"].Value?.ToString()??string.Empty;
                string valueK52 = ws1.Range["K52"].Value?.ToString()??string.Empty;
                string combinedString = valueF51 + valueK52;
                string pdfDirectory = System.IO.Path.GetDirectoryName(filePath);
                int sheetCount = workBook.Sheets.Count;
                string sheetName = $"SHEET1TO{sheetCount}";
                string pdfFilePath = System.IO.Path.Combine(pdfDirectory, combinedString + "-" + sheetName + "-CL" + ".pdf");
                workBook.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, pdfFilePath);
                // Close the workbook and Excel application
                workBook.Close(false);
                excelApp.Quit();

                // Release COM objects
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

                // Re-enable the button and update the text after export is complete
                generatebtn.Enabled = true;
                //generatebtn.Text = "Generate To PDF";

                // Optionally, you could display a MessageBox, or show a toast-style notification.
                MessageBox.Show("Excel data exported to PDF successfully!");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}");
            }
        }

        private void ExportToPDF_Load(object sender, EventArgs e)
        {

        }
    }

}
