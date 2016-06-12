using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;

namespace Com.Ericmas001.Office.Excel
{
    public static class ExcelUtil
    {
        public static IEnumerable<string> GetSheetNames(string excelFile)
        {
            OleDbConnection objConn = null;
            DataTable dt = null;

            try
            {
                objConn = new OleDbConnection($"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={excelFile};Extended Properties=\"Excel 12.0 Xml;HDR=YES;\"");
                objConn.Open();
                dt = objConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                return dt == null ? null : (from DataRow row in dt.Rows select row["TABLE_NAME"].ToString() into name where name.EndsWith("$") select name.Remove(name.Length - 1)).ToList();
            }
            catch
            {
                return null;
            }
            finally
            {
                if (objConn != null)
                {
                    objConn.Close();
                    objConn.Dispose();
                }
                dt?.Dispose();
            }
        }
        public static DataTable GetSheetData(string excelFile, string sheetname)
        {
            try
            {
                var excelConnection = new OleDbConnection($"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={excelFile};Extended Properties=\"Excel 12.0 Xml;HDR=YES;\"");
                excelConnection.Open();
                
                var dbCommand = new OleDbCommand($"SELECT * FROM [{sheetname}$]", excelConnection);
                var dataAdapter = new OleDbDataAdapter(dbCommand);
                var dTable = new DataTable();
                dataAdapter.Fill(dTable);

                dTable.Dispose();
                dataAdapter.Dispose();
                dbCommand.Dispose();
                excelConnection.Close();
                excelConnection.Dispose();

                return dTable;
            }
            catch
            {
                return null;
            }
        }
    }
}
