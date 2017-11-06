using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data;

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

    }
}
