using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data;
using System.Data.OleDb;
using ADODB;

namespace fUtility
{
    public static class DatabaseFunctionality
    {
        public static DataTable GetDataTableFromSql(string sql, string connectionString)
        {
            SqlConnection conn = new SqlConnection();
            conn.ConnectionString = connectionString;
            conn.Open();


            SqlCommand sqlQuery = new SqlCommand(sql);
            sqlQuery.Connection = conn;
            SqlDataAdapter adapter = new SqlDataAdapter(sqlQuery);

            var table = new DataTable();
            adapter.Fill(table);

            conn.Close();
            return table;
        }

        public static DataTable GetDataTableFromSQLOleDb(string sql, string connectionString)
        {
            OleDbConnection conn = new OleDbConnection();
            conn.ConnectionString = connectionString;
            conn.Open();

            OleDbDataAdapter adapter = new OleDbDataAdapter(sql, conn);
            DataTable table = new DataTable();
            adapter.Fill(table);

            conn.Close();
            return table;
        }

        public static DataTable GetDataTableFromSQLAdoDb(string sql, string connectionString)
        {
            var conn = new ADODB.Connection();
            var recordSet = new ADODB.Recordset();
            conn.Open(connectionString);
            recordSet.Open(sql, conn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, 0);

            var table = new DataTable();
            var adapter = new OleDbDataAdapter();
            adapter.Fill(table, recordSet);

            conn.Close();
            return table;

        }

    }
}
