using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.Odbc;

namespace fUtility
{
    public static class QueryCSV
    {


        public static DataTable QueryCSVAndStoreAsDataTable(string sql)
        {
            string connectionString = @"Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=C:;Extensions=csv,txt";
            var connection = new OdbcConnection(connectionString);
            connection.Open();

            var adapter = new OdbcDataAdapter(sql, connection);

            DataTable dataTable = new DataTable();
            adapter.Fill(dataTable);
            return dataTable;
        }

        public static DataTable QueryDataTable(DataTable table, string sql, string folder)
        {
            string tableName = table.TableName;
            string fullPath = folder + "\\" + tableName + ".CSV";
            sql = sql.Replace(tableName, fullPath);

            table.WriteToCsvFile(fullPath);
            return QueryCSVAndStoreAsDataTable(sql);
        }

        public static DataTable QueryDataTable(List<DataTable> tables, string sql, string folder)
        {
            string[] tableNames = tables.Select(s => s.TableName).ToArray();
            string[] fullPaths = tableNames.Select(s => folder + "\\" + s + ".CSV").ToArray();

            for (int i = 0; i<tables.Count; i++)
            {
                tables[i].WriteToCsvFile(fullPaths[i]);
                sql = sql.Replace(tableNames[i], fullPaths[i]);
            }

            return QueryCSVAndStoreAsDataTable(sql);


        }
    }
}
