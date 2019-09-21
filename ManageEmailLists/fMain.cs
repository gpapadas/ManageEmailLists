using System;
using System.Collections.Generic;
using System.IO;
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
        private string fileName, newFileName, fileExtension;
        private List<string> emails;

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
                // Get file's extension.
                FileInfo fileInfo = new FileInfo(fileName);
                fileExtension = fileInfo.Extension;
            }

            Excel.Application excelApp = new Excel.Application();

            // Read the excel file.
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

            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWorkbook;
            Excel.Worksheet excelWorksheet;
            object missingValue = System.Reflection.Missing.Value;

            excelWorkbook = excelApp.Workbooks.Add(missingValue);

            excelWorksheet = (Excel.Worksheet)excelWorkbook.Worksheets.get_Item(1);
            excelWorksheet.Cells[1, 1] = "Email";

            for (int index = 0; index < removedDuplicates.Count; index++)
            {
                excelWorksheet.Cells[index + 2, 1] = removedDuplicates[index];
            }

            // Check the file extension and save the new file accordingly.
            if (fileExtension == ".xlsx")
            {
                newFileName = fileName.Replace(".xlsx", " - without duplicates.xlsx");
                excelWorkbook.SaveAs(newFileName, Excel.XlFileFormat.xlOpenXMLWorkbook, missingValue, missingValue,
                    false, false, Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlUserResolution, true,
                    missingValue, missingValue, missingValue);
            }
            else
            {
                newFileName = fileName.Replace(".xls", " - without duplicates.xls");
                excelWorkbook.SaveAs(newFileName, Excel.XlFileFormat.xlWorkbookNormal, missingValue, missingValue,
                    missingValue, missingValue, Excel.XlSaveAsAccessMode.xlExclusive, missingValue, missingValue,
                    missingValue, missingValue, missingValue);
            }

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
            List<string> personalEmails = new List<string>();
            List<string> businessEmails = new List<string>();
            int index = 0;

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
            }

            // Export business emails.
            for (index = 0; index < businessEmails.Count; index++)
            {
                excelWorkSheet.Cells[index + 2, 4] = businessEmails[index];
            }

            // Export websites from the business emails.
            for (index = 0; index < businessEmails.Count; index++)
            {
                excelWorkSheet.Cells[index + 2, 8] =
                    businessEmails[index].Substring(businessEmails[index].IndexOf("@") + 1,
                    (businessEmails[index].Length - 1) - businessEmails[index].IndexOf("@"));
            }

            // Check the file extension and save the new file accordingly.
            if (fileExtension == ".xlsx")
            {
                newFileName = fileName.Replace(".xlsx", " - personal and business emails.xlsx");
                excelWorkbook.SaveAs(newFileName, Excel.XlFileFormat.xlOpenXMLWorkbook, missingValue, missingValue,
                    false, false, Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlUserResolution, true,
                    missingValue, missingValue, missingValue);
            }
            else
            {
                newFileName = fileName.Replace(".xls", " - personal and business emails.xls");
                excelWorkbook.SaveAs(newFileName, Excel.XlFileFormat.xlWorkbookNormal, missingValue, missingValue,
                    missingValue, missingValue, Excel.XlSaveAsAccessMode.xlExclusive, missingValue, missingValue,
                    missingValue, missingValue, missingValue);
            }

            excelWorkbook.Close(true, missingValue, missingValue);
            excelApp.Quit();

            ReleaseObject(excelWorkSheet);
            ReleaseObject(excelWorkbook);
            ReleaseObject(excelApp);

            MessageBox.Show("The excel file created. You can find it at: " + newFileName, "File created",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
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

        /// <summary>
        /// Removes all duplicates from a list.
        /// </summary>
        /// <param name="inputList"></param>
        /// <returns></returns>
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

        /// <summary>
        /// Releases an object.
        /// </summary>
        /// <param name="o"></param>
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
