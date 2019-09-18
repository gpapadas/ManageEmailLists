using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Data.OleDb;
using System.Data.Common;
using System.Windows.Forms;

namespace ManageEmailLists
{
    public partial class fMain : Form
    {
        private string connectionString;
        private string filename, newFilename;
        private List<string> initialList;
        string[] sheetNames;

        public fMain()
        {
            InitializeComponent();
        }

        private void BtnImportExcel_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();

            dialog.Filter = "ms excel 2003 files (*.xls)|*.xls|ms excel 2007 files (*.xlsx)|*.xlsx";
            dialog.InitialDirectory = "C:";
            dialog.Title = "Άνοιγμα αρχείου excel";
            if (dialog.ShowDialog() == DialogResult.OK)
                filename = dialog.FileName;

            initialList = new List<string>();

            sheetNames = GetExcelSheetNames(filename);
            connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;
                Data Source=" + filename + ";Extended Properties=Excel 8.0;";

            DbProviderFactory factory = DbProviderFactories.GetFactory("System.Data.OleDb");

            using (DbConnection connection = factory.CreateConnection())
            {
                connection.ConnectionString = connectionString;

                using (DbCommand command = connection.CreateCommand())
                {
                    command.CommandText = "SELECT Email FROM [" + sheetNames[0] + "]";

                    connection.Open();

                    using (DbDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            initialList.Add(reader["Email"].ToString());
                        }
                    }
                    connection.Close();
                }
            }

            btnFindDuplicates.Enabled = true;
            btnExportEmailsWithoutDuplicates.Enabled = true;
            btnExportPrivateCorporateEmails.Enabled = true;
        }

        private void BtnExportEmailsWithoutDuplicates_Click(object sender, EventArgs e)
        {

        }

        private void BtnExportPrivateCorporateEmails_Click(object sender, EventArgs e)
        {

        }

        private String[] GetExcelSheetNames(string excelFile)
        {
            OleDbConnection connection = null;
            System.Data.DataTable dt = null;

            try
            {
                // Create connection string.
                //string connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;" +
                //    "Data Source=" + filename + ";Extended Properties=Excel 8.0;";
                string connectionString =
                    String.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=Excel 8.0;", excelFile);
                //string connectionString =
                //    String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 8.0;HDR=YES\";", excelFile);

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
    }
}
