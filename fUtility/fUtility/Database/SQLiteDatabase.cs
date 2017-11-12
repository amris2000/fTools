using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SQLite;
using System.Data;
using System.Collections.ObjectModel;

namespace fUtility
{
    public class SQLiteDatabase
    {
        public SQLiteConnection Conn;
        private SQLiteCommand _cmd;
        private SQLiteDataAdapter _adapter;
        private string _dbName;

        private void SetConnection()
        {
            Conn = new SQLiteConnection("Data Source=" + _dbName + ";Version=3;New=False;Compress=True;");
        }

        public void QueryTable(string query)
        {
            SetConnection();
            Conn.Open();
            _cmd = Conn.CreateCommand();
            _cmd.CommandText = query;
            _cmd.ExecuteNonQuery();
            Conn.Close();
        }

        public void CreateTable(string tableName, DataTable table)
        {
            //Conn.Open();

            // This ensures that culture is not DA in which case ToString on doubles converts to "2,22235" 
            // and this fucks up the insert query... 
            System.Globalization.CultureInfo.DefaultThreadCurrentCulture = System.Globalization.CultureInfo.InvariantCulture;

            // Create table with appropiate headers and data types
            // Try to delete
            string sqlDelete = "DROP TABLE IF EXISTS " + tableName;
            _cmd = new SQLiteCommand(sqlDelete, Conn);
            _cmd.ExecuteNonQuery();

            string sql = CreateTableQuery(tableName, table);
            _cmd = new SQLiteCommand(sql, Conn);
            _cmd.ExecuteNonQuery();

            List<string> columnNames = new List<string>();
            foreach (DataColumn col in table.Columns)
                columnNames.Add(col.ColumnName);

            // Fill rows of table with rows from dataTable
            string insertQuery;

            // Slow, execute all rows from one ExecuteNonQuery call...
            foreach (DataRow row in table.Rows)
            {
                insertQuery = CreateInsertRowQuery(tableName, row, columnNames);
                _cmd = new SQLiteCommand(insertQuery, Conn);
                _cmd.ExecuteNonQuery();
            }

            //Conn.Close();
        }

        public SQLiteDatabase(string databaseName)
        {
            _dbName = databaseName;
            SetConnection();
            SQLiteConnection.CreateFile(_dbName);
            Conn.Open();

        }

        public DataTable GetDataTable(string tableName)
        {
            //Conn.Open();
            _adapter = new SQLiteDataAdapter(String.Format("SELECT * FROM {0}", tableName),Conn);
            var table = new DataTable();
            _adapter.Fill(table);
            //Conn.Close();
            return table;
        }

        public DataTable InspectDatabase()
        {
            //Conn.Open();
            _adapter = new SQLiteDataAdapter("SELECT * FROM sqlite_master", Conn);

            var dataTable = new DataTable();
            _adapter.Fill(dataTable);
            //Conn.Close();
            return dataTable;

        }

        private string CreateTableQuery(string tableName, DataTable table)
        {
            //List<string> columnTypes = new List<string>();
            //List<string> columnNames = new List<string>();

            //foreach (DataColumn col in table.Columns)
            //{
            //    columnTypes.Add(DataTableTypeToSqlConversion(col));
            //    columnNames.Add(col.ColumnName);
            //}

            //for (int i = 0; i < columnTypes.Count; i++)
            //    output = output + columnNames[i] + " " + columnTypes[i] + ",";

            //output = output.Substring(0, output.Length - 1); // Remove last ","

            List<string> ColumnNamesAndTypes = new List<string>();

            foreach (DataColumn col in table.Columns)
                ColumnNamesAndTypes.Add(col.ColumnName + " " + DataTableTypeToSqlConversion(col));

            string output = String.Join(",", ColumnNamesAndTypes.ToArray());
            return "CREATE TABLE " + tableName + "( " + output + ")";

        }

        private string CreateInsertRowQuery(string tableName, DataRow row, List<string> ColumnNames)
        {
            // Warning for this method @ https://stackoverflow.com/questions/19479166/sqlite-simple-insert-query
            // Probably related to SQL injection

            string rowAsStr;
            string temp;

            if (row.ItemArray[0].ToString().Contains("'"))
                temp = row.ItemArray[0].ToString().Replace("'", "''");
            else
                temp = row.ItemArray[0].ToString();

            if (row.ItemArray[0] is string)
            {
                rowAsStr = "'" + temp + "'";
            }
            else
                rowAsStr = temp;

            for (int i = 1; i < row.ItemArray.Length; i++)
            {
                if (row.ItemArray[i].ToString().Contains("'"))
                    temp = row.ItemArray[i].ToString().Replace("'", "''");
                else
                    temp = row.ItemArray[i].ToString();

                if (row.ItemArray[i] is string)
                {
                    rowAsStr = rowAsStr + ",'" + temp + "'";
                }
                else
                    rowAsStr = rowAsStr + ", " + row.ItemArray[i].ToString();
            }

            string columnNamesAsStr = String.Join(",", ColumnNames.ToArray());

            //foreach (string colName in ColumnNames)
            //    columnNamesAsStr = columnNamesAsStr + colName + ", ";

            //columnNamesAsStr = columnNamesAsStr.Substring(0,columnNamesAsStr.Length - 2);

            return "INSERT INTO " + tableName + " (" + columnNamesAsStr + ") VALUES (" + rowAsStr + ")";
        }

        private string DataTableTypeToSqlConversion(DataColumn column)
        {
            string output = "NONE"; // Default, undefind data type in SQLite

            //var typeSwitch = new TypeSwitch()
            //    .Case((int x) => output = "INTEGER")
            //    .Case((bool x) => output = "BOOL")
            //    .Case((double x) => output = "NUMERIC")
            //    .Case((System.String x) => output = "TEXT");

            //typeSwitch.Switch(column.DataType);

            var @switch = new Dictionary<Type, Action>
            {
                { typeof(System.String), () => output = "TEXT" },
                { typeof(System.Int32), () => output = "INTEGER"},
                { typeof(System.Double), () => output = "REAL" },
                { typeof(System.DateTime), () => output = "TEXT" },
            };

            @switch[column.DataType]();

            return output;
        }

        public DataTable QueryTableAndReturnResult(string query)
        {
            SetConnection();
            Conn.Open();
            _cmd = Conn.CreateCommand();
            _adapter = new SQLiteDataAdapter(query, Conn);
            var table = new DataTable();
            _adapter.AcceptChangesDuringFill = false;
            _adapter.Fill(table);
            Conn.Close();

            return table;
        }

    }
}
