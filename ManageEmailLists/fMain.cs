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
        private string fileName, newFileName;
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

            dialog.Filter = "Ms Excel 2007 files (*.xlsx)|*.xlsx|Ms Excel 2003 files (*.xls)|*.xls";
            dialog.InitialDirectory = "C:";
            dialog.Title = "Open excel file";

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                fileName = dialog.FileName;
            }

            Excel.Application excelApp = new Excel.Application();

            if (excelApp != null)
            {
                Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(fileName, 0, true, 5, "", "", true,
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

                        // Add to list.
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
            List<string> duplicates = FindDuplicates(emails);

            lbDuplicateEmails.Items.Clear();
            foreach (string value in duplicates)
            {
                lbDuplicateEmails.Items.Add(value);
            }
        }

        private void BtnExportEmailsWithoutDuplicates_Click(object sender, EventArgs e)
        {
            List<string> removedDuplicates = RemoveDuplicates(emails);

            //Excel.Application xlApp;
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWorkbook;
            Excel.Worksheet excelWorksheet;
            object missingValue = System.Reflection.Missing.Value;

            //xlApp = new Excel.ApplicationClass();
            excelWorkbook = excelApp.Workbooks.Add(missingValue);

            excelWorksheet = (Excel.Worksheet)excelWorkbook.Worksheets.get_Item(1);
            excelWorksheet.Cells[1, 1] = "Email";

            for (int index = 0; index < removedDuplicates.Count; index++)
            {
                excelWorksheet.Cells[index + 2, 1] = removedDuplicates[index];
            }

            newFileName = fileName.Replace(".xlsx", " - without duplicates.xlsx");
            excelWorkbook.SaveAs(newFileName, Excel.XlFileFormat.xlOpenXMLWorkbook, missingValue, missingValue,
                false, false, Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlUserResolution, true,
                missingValue, missingValue, missingValue);

            //TODO: Commented code below is for xls files. I need to check the version of Excel (xls, xlsx)

            //excelWorkbook.SaveAs(newFileName, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue,
            //    misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);

            excelWorkbook.Close(true, missingValue, missingValue);
            excelApp.Quit();

            // Release objects.
            ReleaseObject(excelWorksheet);
            ReleaseObject(excelWorkbook);
            ReleaseObject(excelApp);

            MessageBox.Show("The excel file created. You can find it at: " + newFileName, "File created",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void BtnExportPersonalBusinessEmails_Click(object sender, EventArgs e)
        {
            List<string> gmailList = new List<string>();
            List<string> yahooList = new List<string>();
            List<string> hotmailList = new List<string>();
            List<string> windowsLiveList = new List<string>();
            //List<string> otenetList = new List<string>();
            List<string> personalEmails = new List<string>();
            List<string> businessEmails = new List<string>();
            int index, outIndex = 0;

            foreach (string email in emails)
            {
                if (email.Contains("@gmail") || email.Contains("@yahoo")
                    || email.Contains("@hotmail") || email.Contains("@windowslive"))
                {
                    personalEmails.Add(email);
                }
                else
                {
                    businessEmails.Add(email);
                }
            }

            Excel.Application excelApp;
            Excel.Workbook excelWorkbook;
            Excel.Worksheet excelWorkSheet;
            object missingValue = System.Reflection.Missing.Value;

            excelApp = new Excel.Application();
            excelWorkbook = excelApp.Workbooks.Add(missingValue);

            excelWorkSheet = (Excel.Worksheet)excelWorkbook.Worksheets.get_Item(1);
            excelWorkSheet.Cells[1, 1] = "Personal emails";
            excelWorkSheet.Cells[1, 4] = "Business emails";
            excelWorkSheet.Cells[1, 8] = "Website of business emails";

            // Export personal emails.
            for (index = 0; index < personalEmails.Count; index++)
            {
                excelWorkSheet.Cells[index + 2, 1] = personalEmails[index];
                //outIndex++;
            }

            // Export business emails.
            for (index = 0; index < businessEmails.Count; index++)
            {
                excelWorkSheet.Cells[index + 2, 4] = businessEmails[index];
                //outIndex++;
            }

            // Export websites from the business emails.
            for (index = 0; index < businessEmails.Count; index++)
            {
                excelWorkSheet.Cells[index + 2, 8] =
                    businessEmails[index].Substring(businessEmails[index].IndexOf("@") + 1,
                    (businessEmails[index].Length - 1) - businessEmails[index].IndexOf("@"));
            }

            newFileName = fileName.Replace(".xlsx", "- personal and private emails.xlsx");
            excelWorkbook.SaveAs(newFileName, Excel.XlFileFormat.xlOpenXMLWorkbook, missingValue, missingValue,
                false, false, Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlUserResolution, true,
                missingValue, missingValue, missingValue);

            //TODO: Commented code below is for xls files. I need to check the version of Excel (xls, xlsx)

            //excelWorkbook.SaveAs(newFileName, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue,
            //    misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            excelWorkbook.Close(true, missingValue, missingValue);
            excelApp.Quit();

            ReleaseObject(excelWorkSheet);
            ReleaseObject(excelWorkbook);
            ReleaseObject(excelApp);

            MessageBox.Show("The excel file created. You can find it at: " + newFileName, "File created",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private String[] GetExcelSheetNames(string excelFile)
        {
            //TODO: Delete all the method probably.

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

        /// <summary>
        /// Finds duplicates in a list and returns its distinct values.
        /// </summary>
        /// <param name="inputList"></param>
        /// <returns></returns>
        private static List<string> FindDuplicates(List<string> inputList)
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

        private List<string> RemoveDuplicates(List<string> inputList)
        {

            Dictionary<string, int> dict = new Dictionary<string, int>();
            List<string> outputList = new List<string>();

            foreach (string input in inputList)
            {
                if (!dict.ContainsKey(input))
                {
                    dict.Add(input, 0);
                    outputList.Add(input);
                }
            }
            return outputList;
        }

        private void ReleaseObject(object o)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(o);
                o = null;
            }
            catch (Exception e)
            {
                o = null;
                MessageBox.Show("Exception Occured while releasing object " + e.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
