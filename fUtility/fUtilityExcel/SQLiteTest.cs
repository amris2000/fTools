using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using System.Data;
using fUtility;

namespace fUtilityExcel
{
    public class SQLiteTest
    {
        [ExcelFunction(Name = Constant.FunctionPrefix + "SQL.DataTable.Put", IsVolatile = true)]
        public static string fUtilityPutDataTable(string handle, object[,] data)
        {
            PersistentObjects.AssertReservedWordsTableName(handle);
            var table = ExcelFriendlyConversion.ConvertObjectArrayToDataTable(handle, data);
            PersistentObjects.MyDatabase.CreateTable(handle, table);
            return handle;
        }

        // NEEDS TO BE IMPLEMENTED ....
        [ExcelFunction(Name = Constant.FunctionPrefix + "SQL.DataTable.Query", IsVolatile = true)]
        public static object[,] fUtilityQueryDataTable(string sqlStringOrHandle, object noHeadersInput)
        {
            string sql = ObjectInfo.DetermineIfSqlHandleOrSqlQuery(sqlStringOrHandle);

            object[,] output;
            bool noHeaders = Optional.Check(noHeadersInput, false);
            var tableHandle = ParsingFunctionality.GetTableNameFromSqlQuery(sql);

            if (PersistentObjects.ContainsKey(tableHandle, "DATATABLE") == false)
            {
                string errorMessage = "ERROR: DataTable + " + tableHandle + " does not exist in the object map.";
                output = new object[1, 1];
                output[0, 0] = errorMessage;
                return output;
            }
            else
            {
                var newTable = QueryCSV.QueryDataTable(sql, PersistentObjects.CsvPath);
                output = ExcelFriendlyConversion.ConvertDataTableToObjectArray(newTable, noHeaders);
                return output;
            }
        }

        [ExcelFunction(Name = Constant.FunctionPrefix + "SQL.DataTable.PutFromSQL", IsVolatile = true)]
        public static string fUtilityDataBasePutDataTableFromSQL(string handle, string connectionString, string sqlStringOrHandle)
        {
            PersistentObjects.AssertReservedWordsTableName(handle);

            string sql = ObjectInfo.DetermineIfSqlHandleOrSqlQuery(sqlStringOrHandle);

            if (ExcelDnaUtil.IsInFunctionWizard())
                return "no wizard.";
            var table = DatabaseFunctionality.GetDataTableFromSql(sql, connectionString);

            table.TableName = handle;
            PersistentObjects.MyDatabase.CreateTable(handle, table);
            return handle;
        }


    }
}
