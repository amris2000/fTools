using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SQLite;

namespace fUtility
{
    public static class PersistentObjects
    {
        public const string DEFAULT_FIELD_NAME = "DEFAULT";
        public static string CsvPath { get; private set; }
        private static bool _csvPathIsSet; 

        public static void SetDefaultCsvPath(string path)
        {
            CsvPath = path;
            _csvPathIsSet = true;
        }

        //public static SQLiteDatabase MyDatabase;

        static PersistentObjects()
        {
            //MyDatabase = new SQLiteDatabase("FTDB");
        }


        public static void AssertCsvPathIsSet()
        {
            if (_csvPathIsSet == false)
                throw new InvalidOperationException("CSV path has not been set.");
        }

        public static void AssertReservedWordsTableName(string tableName)
        {
            string[] reservedWords = new string[] { "SELECT", "FROM", "*", "JOIN", "LEFT", "ORDER", "BY", "SORT" };
            if (reservedWords.ToList<string>().Contains(tableName.ToUpper()))
                throw new InvalidOperationException("TABLENAME CANNOT BE EQUAL TO " + tableName.ToUpper());
        }

        public static IDictionary<string, ObjectStoreEntry> ObjectStore = new Dictionary<string, ObjectStoreEntry>();

        public static string ParseSqlStringToCsvStringAndStoreTables(string sql)
        {
            AssertCsvPathIsSet();

            var array = sql.Split(' ').ToList<string>();
            string tableFullName;
            var newSql = sql.ToUpper();

            // Should consider putting tables in CSV path as soon as they are generated.
            // Multilple queries on large tables will have to put a lot of CSV files..

            foreach (string str in array)
            {
                if (ObjectStore.ContainsKey("DATATABLE" + "." + str))
                {
                    tableFullName = CsvPath + "\\" + str + ".CSV";
                    GetFromMap<DataTable>(str,"DATATABLE").WriteToCsvFile(tableFullName);
                    newSql = newSql.Replace(str, tableFullName);
                }
            }

            return newSql;
        }

        public static bool ContainsKey(string key, string storeField = DEFAULT_FIELD_NAME)
        {
            if (ObjectStore.ContainsKey(storeField.ToUpper() + "." + key.ToUpper()))
                return true;
            else
                return false;
        }

        public static void AddToMap(string handle, object data, string field = null)
        {
            if (String.IsNullOrWhiteSpace(handle))
                throw new InvalidOperationException("ERROR: HANDLE CANNOT BE AN EMPTY STRING");

            ObjectStoreEntry entry = new ObjectStoreEntry(handle, data, field);
            string fullKey = entry.FullKey;

            if (ObjectStore.Keys.Contains(fullKey))
                ObjectStore.Remove(fullKey);

            PersistentObjects.ObjectStore.Add(fullKey, entry);
        }

        public static T GetFromMap<T>(string key, string field = DEFAULT_FIELD_NAME)
        {
            string fullKey = field.ToUpper() + "." + key.ToUpper();

            if (PersistentObjects.ContainsKey(key, field) == false)
                throw new InvalidOperationException("MAP DOES NOT CONTAIN KEY '" + key + "' in field '" + field + "'.");
            else
                return (T)PersistentObjects.ObjectStore[fullKey].Obj;
        }


    }
}
