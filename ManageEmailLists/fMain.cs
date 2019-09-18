using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.OleDb;
using System.Data.Common;
using System.Windows.Forms;

namespace ManageEmailLists
{
    public partial class fMain : Form
    {
        private string connectionString;
        private string filename, newFilename;
        private string path;
        private List<string> initialList;
        private List<string> emails;
        string[] sheetNames;

        public fMain()
        {
            InitializeComponent();
        }

        private void BtnImportExcel_Click(object sender, EventArgs e)
        {
            // Open the file.
            OpenFileDialog dialog = new OpenFileDialog();

            dialog.Filter = "Ms excel 2003 files (*.xls)|*.xls|Ms excel 2007 files (*.xlsx)|*.xlsx";
            dialog.InitialDirectory = "C:";
            dialog.Title = "Open excel file";

            if (dialog.ShowDialog() == DialogResult.OK)
                path = dialog.FileName;

            // Read excel file.
            Excel.Application excelApp = new Excel.Application();

            if (excelApp != null)
            {
                Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(path, 0, true, 5, "", "", true,
                    Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelWorkbook.Sheets[1];
                Excel.Range excelRange = excelWorksheet.UsedRange;

                // Get rows and columns count.
                int rows = excelRange.Rows.Count;
                int cols = excelRange.Columns.Count;

                // Initialize the emails list.
                emails = new List<string>();

                for (int i = 1; i <= rows; i++)
                {
                    for (int j = 1; j <= cols; j++)
                    {
                        Excel.Range range = (excelWorksheet.Cells[i, 1] as Excel.Range);
                        string cellValue = range.Value.ToString();

                        // Add to list
                        emails.Add(cellValue);
                    }
                }

                excelWorkbook.Close();
                excelApp.Quit();
            }

            // Enable the other buttons.
            btnFindDuplicates.Enabled = true;
            btnExportEmailsWithoutDuplicates.Enabled = true;
            btnExportPrivateBusinessEmails.Enabled = true;
        }

        private void BtnFindDuplicates_Click(object sender, EventArgs e)
        {
            List<string> duplicates = findDuplicates(emails);

            lbDuplicateEmails.Items.Clear();
            foreach (string value in duplicates)
            {
                lbDuplicateEmails.Items.Add(value);
            }
        }

        private void BtnExportEmailsWithoutDuplicates_Click(object sender, EventArgs e)
        {

        }

        private void BtnExportPrivateBusinessEmails_Click(object sender, EventArgs e)
        {

        }

        private String[] GetExcelSheetNames(string excelFile)
        {
            //TODO: Delete method.
            OleDbConnection connection = null;
            System.Data.DataTable dt = null;

            try
            {
                // Create connection string.
                string connectionString =
                    String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=Excel 12.0;", excelFile);

                // Create connection object by using the preceding connection string.
                connection = new OleDbConnection(connectionString);
                // Open connection with the database.
                connection.Open();

                // Get the data table containg the schema guid.
                dt = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                if (dt == null)
                {
                    return null;
                }

                String[] excelSheets = new String[dt.Rows.Count];
                int i = 0;

                // Add the sheet name to the string array.
                foreach (DataRow row in dt.Rows)
                {
                    excelSheets[i] = row["TABLE_NAME"].ToString();
                    i++;
                }

                // Loop through all of the sheets if you want too...
                for (int j = 0; j < excelSheets.Length; j++)
                {
                    // Query each excel sheet.
                }

                return excelSheets;
            }
            catch (Exception e)
            {
                throw new Exception("Exception: ", e);
            }
            finally
            {
                // Clean up.
                if (connection != null)
                {
                    connection.Close();
                    connection.Dispose();
                }
                if (dt != null)
                {
                    dt.Dispose();
                }
            }
        }

        private static List<string> findDuplicates(List<string> inputList)
        {
            List<string> duplicates = new List<string>();
            HashSet<string> uniques = new HashSet<string>();

            foreach (var input in inputList)
            {
                if (uniques.Contains(input))
                {
                    duplicates.Add(input);
                }
                else
                {
                    uniques.Add(input);
                }
            }

            return duplicates.Distinct().ToList();
        }
    }
}
