using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Reflection;
using System.Data.OleDb;
//using DocumentFormat.OpenXml.Packaging;
//using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace ModelGenerator
{
    public partial class Form1 : Form
    {
        DataSet ds = new DataSet();
        OleDbDataAdapter adapter = new OleDbDataAdapter();
        List<String> tables = new List<String>();

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog of = new OpenFileDialog();
            of.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";

            if (of.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    string ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + of.InitialDirectory + of.FileName
                        + @";Extended Properties=""Excel 12.0 Macro;HDR=Yes;ImportMixedTypes=Text;TypeGuessRows=0""";

                    ds = ReadExcelFile(ConnectionString);
                    comboBox1.DataSource = tables;
                    dataGridView1.DataSource = ds.Tables[0];
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not read file from disk. Original error: " + ex.Message);
                }
            }
        }
        private DataSet ReadExcelFile(string connectionString)
        {
            DataSet ds = new DataSet();

            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = conn;

                // Get all Sheets in Excel File
                DataTable dtSheet = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                // Loop through all Sheets to get data
                foreach (DataRow dr in dtSheet.Rows)
                {
                    string sheetName = dr["TABLE_NAME"].ToString();

                    if (!sheetName.EndsWith("$"))
                        continue;

                    // Get all rows from the Sheet
                    cmd.CommandText = "SELECT * FROM [" + sheetName + "]";

                    DataTable dt = new DataTable();
                    dt.TableName = sheetName;

                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                    da.Fill(dt);

                    tables.Add(dt.TableName);
                    ds.Tables.Add(dt);
                }

                cmd = null;
                conn.Close();
            }

            return ds;
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            String selection = comboBox1.SelectedItem as String;
            var index = tables.FindIndex(a => a == selection);
            dataGridView1.DataSource = ds.Tables[index];
        }
    }
}
