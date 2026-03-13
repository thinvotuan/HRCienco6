using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Web;

namespace BatDongSan.Helper.Utils
{
    public class ExcelDataProcessing
    {
        /// <summary>
        /// Gets or sets the Excel filename
        /// </summary>
        private string FileName { get; set; }

        /// <summary>
        /// Template connectionstring for Excel connections
        /// </summary>
        private const string ConnectionStringTemplate = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=Excel 8.0;";

        /// <summary>
        /// Default constructor
        /// </summary>
        /// <param name="fileName">The Excel file to process</param>
        public ExcelDataProcessing(string fileName)
        {
            FileName = fileName;
        }

        /// <summary>
        /// Returns a worksheet as a linq-queryable enumeration
        /// </summary>
        /// <param name="sheetName">The name of the worksheet</param>
        /// <returns>An enumerable collection of the worksheet</returns>
        public EnumerableRowCollection<DataRow> GetWorkSheet(string sheetName)
        {
            // Build the connectionstring
            string connectionString = OpenExcelFile(FileName);// string.Format(ConnectionStringTemplate, FileName);

            // Query the specified worksheet
            OleDbDataAdapter dataAdapter = new OleDbDataAdapter(string.Format("SELECT * FROM [{0}$]", sheetName), connectionString);

            // Fill the dataset from the data adapter
            DataSet myDataSet = new DataSet();
            dataAdapter.Fill(myDataSet, "ExcelInfo");

            // Initialize a data table which we can use to enumerate the contents based on the dataset
            DataTable dataTable = myDataSet.Tables["ExcelInfo"];

            // Return the data table contents as a queryable enumeration
            return dataTable.AsEnumerable();
        }

        public DataTable GetDataTableWorkSheet(string sheetName)
        {
            // Build the connectionstring
            string connectionString = OpenExcelFile(FileName);// string.Format(ConnectionStringTemplate, FileName);

            // Query the specified worksheet
            OleDbDataAdapter dataAdapter = new OleDbDataAdapter(string.Format("SELECT * FROM [{0}$]", sheetName), connectionString);

            // Fill the dataset from the data adapter
            DataSet myDataSet = new DataSet();
            dataAdapter.Fill(myDataSet, "ExcelInfo");

            // Initialize a data table which we can use to enumerate the contents based on the dataset
            DataTable dataTable = myDataSet.Tables["ExcelInfo"];

            return dataTable;
        }

        public string SaveDataToExcel(string sql)
        {
            // Build the connectionstring
            string connectionString = string.Format(ConnectionStringTemplate, FileName);
            try
            {
                System.Data.OleDb.OleDbConnection MyConnection;
                System.Data.OleDb.OleDbCommand myCommand = new System.Data.OleDb.OleDbCommand();
                MyConnection = new System.Data.OleDb.OleDbConnection(connectionString);
                MyConnection.Open();
                myCommand.Connection = MyConnection;
                myCommand.CommandText = sql;
                myCommand.ExecuteNonQuery();
                MyConnection.Close();

                return null;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        private string OpenExcelFile(string fPath)
        {
            string connectionstring = String.Empty;
            string[] splitdot = fPath.Split(new char[1] { '.' });
            string dot = splitdot[splitdot.Length - 1].ToLower();
            if (dot == "xls")
            {
                //tao chuoi ket noi voi Excel 2003
                connectionstring = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fPath + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=1;HTML Import\"";
            }
            else if (dot == "xlsx")
            {
                //tao chuoi ket noi voi Excel 2007
                connectionstring = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fPath + ";Extended Properties=\"Excel 12.0;HDR=YES;IMEX=1;HTML Import\"";
            }
            return connectionstring;
        }

        /// <summary>
        /// Gets the name of the sheet.
        /// </summary>
        /// <param name="fPath">The f path.</param>
        /// <returns>return all sheet name</returns>
        public ArrayList GetSheetName(string fPath)
        {
            ArrayList sheetnames = new ArrayList();
            string connectionstring = OpenExcelFile(fPath);
            //mo ket noi den file excel
            OleDbConnection cnn = new OleDbConnection(connectionstring);
            cnn.Open();

            //tao bang luu tru tam cac du lieu trong file
            DataTable table = new DataTable();
            table = cnn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            //doc tung dong trong bang luu tru tam
            for (int i = 0; i < table.Rows.Count; i++)
            {
                string name = table.Rows[i][2].ToString().Replace("'", "");//get ten tung sheet co trong bang luu tru
                //kiem tra sheet
                if (name.EndsWith("$"))
                {
                    sheetnames.Add(name.Replace("$", ""));
                }
            }
            cnn.Close();
            table.Dispose();
            return sheetnames;
        }
    }
}