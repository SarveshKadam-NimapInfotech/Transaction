using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Data;

namespace Oledb
{
    internal class Program
    {
        private static object listBox1;

        static void Main(string[] args)
        {
            string connectionString = @"provider = Microsoft.ACE.OLEDB.12.0; 
                            data source = C:\\Users\\Nimap\\Downloads\\backups\\Daily sales - Copy.xlsx; 
                            Extended Properties = 'Excel 12.0'";
            List<string> sheetNames = GetExcelSheetNames(connectionString);

            DataSet dtAllSheets = LoadAllSheetsFromExcel(connectionString, sheetNames);

        }

        public static DataSet LoadAllSheetsFromExcel(string connectionString, List<string> sheetNames)
        {

            DataSet dataSet = new DataSet();
            foreach (string sheetName in sheetNames)
            {
                OleDbConnection oleDbConnection = new OleDbConnection(connectionString);
                DataTable dataTable = new DataTable();
                string sqlQuery = string.Format("SELECT * FROM [{0}]", sheetName);
                oleDbConnection.Open();
                OleDbDataAdapter oleDbDataAdapter = new OleDbDataAdapter(sqlQuery, oleDbConnection);
                oleDbDataAdapter.Fill(dataTable);
                dataSet.Tables.Add(dataTable);
                oleDbConnection.Close();
            }
            return dataSet;

        }

        public static List<string> GetExcelSheetNames(string connectionString)
        {

            OleDbConnection oleDbConnection = new OleDbConnection(connectionString);
            oleDbConnection.Open();
            DataTable dataTable = new DataTable();
            dataTable = oleDbConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            oleDbConnection.Close();
            List<string> sheetNames = new List<string>();
            foreach (DataRow row in dataTable.Rows)
            {
                sheetNames.Add(row["TABLE_NAME"].ToString());
            }
            return sheetNames;

        }

        //private static List<string> GetExcelSheetNames(string connectionString)
        //{
        //    OleDbConnection oleDbConnection = new OleDbConnection(connectionString);
        //    oleDbConnection.Open();
        //    DataTable dataTable = new DataTable();
        //    dataTable = oleDbConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
        //    oleDbConnection.Close();
        //    List<string> sheetNames = new List<string>();
        //    foreach (DataRow row in dataTable.Rows)
        //    {
        //        sheetNames.Add(row["TABLE_NAME"].ToString());
        //    }
        //    return sheetNames;
        //}


        //////////////////////////////////////////////////////////////////////////////////////////
        

        //string excelFilePath = "C:\\Users\\Nimap\\Downloads\\backups\\Daily sales - Copy.xlsx";
        //string sheetName = "[2023$]"; // Replace with your sheet name

        //string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={excelFilePath};Extended Properties='Excel 12.0 Xml;HDR=YES';";

        //string query = "SELECT * FROM " + sheetName + " WHERE [Column3] = ? AND [Column7] LIKE ?";

        //using (OleDbConnection connection = new OleDbConnection(connectionString))
        //{
        //    using (OleDbCommand command = new OleDbCommand(query, connection))
        //    {
        //        connection.Open();
        //        command.Parameters.AddWithValue("@param1", DateTime.Parse("10/10/2023").ToShortDateString());
        //        command.Parameters.AddWithValue("@param2", "D*");

        //        using (OleDbDataAdapter adapter = new OleDbDataAdapter(command))
        //        {
        //            DataSet dataSet = new DataSet();
        //            adapter.Fill(dataSet);

        //            DataTable dataTable = dataSet.Tables[0];

        //            foreach (DataRow row in dataTable.Rows)
        //            {
        //                // Access columns by name or index
        //                string column3Value = row["Column3"].ToString();
        //                string column7Value = row["Column7"].ToString();

        //                Console.WriteLine($"Column3: {column3Value}, Column7: {column7Value}, ...");
        //                // Perform operations with the filtered data
        //            }
        //        }
        //    }
        //}

    }
}
