﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using System.Data;
using fUtility;

namespace fUtilityExcel
{
    public static class Constant
    {
        public const string FunctionPrefix = "ft.";
        public const string DefaultCsvHandle = "AUTOGENERATED_DEFAULT_CSV_HANDLE";
    }

    public class DataTableFunctions
    {
        [ExcelFunction(Name = Constant.FunctionPrefix + "DataTable.GetFields", IsVolatile = true)]
        public static object[,] fUtilityDataTablesGetFields(string handle)
        {
            var table = ObjectMapInterface.GetFromMap<DataTable>(handle);
            var fields = DataTableUtility.GetFields(table);
            var output = new object[fields.Count, 1];
            for (int i = 0; i < fields.Count; i++)
                output[i, 0] = fields[i];

            return output;
        }

        [ExcelFunction(Name = Constant.FunctionPrefix + "DataTable.GetColumnCount", IsVolatile = true)]
        public static int fUtilityDataTablesGetColumnCount(string handle)
        {
            var table = ObjectMapInterface.GetFromMap<DataTable>(handle);
            return table.Columns.Count;
        }

        [ExcelFunction(Name = Constant.FunctionPrefix + "DataTable.GetRowCount", IsVolatile = true)]
        public static int fUtilityDataTablesGetRowCount(string handle)
        {
            var table = ObjectMapInterface.GetFromMap<DataTable>(handle);
            return table.Rows.Count;
        }

        [ExcelFunction(Name = Constant.FunctionPrefix + "ReadSql", IsVolatile = true)]
        public static string fUtilityReadSql(object[] array)
        {
            return ParsingFunctionality.ReadSqlFromArray(DataUtility.removeEmptyEntries(array));
        }

        [ExcelFunction(Name = Constant.FunctionPrefix + "PutSql")]
        public static string fUtilityObjectsStoreSql(string handle, object[] sqlStatement)
        {
            string sql = fUtilityReadSql(sqlStatement);
            ObjectMapInterface.AddToMap(handle, sql);
            return handle;
        }

        [ExcelFunction(Name = Constant.FunctionPrefix + "SetDefaultCsvPath")]
        public static string fUtilitySetDefaultCsvPath(string path)
        {
            ObjectMapInterface.AddToMap(Constant.DefaultCsvHandle, path);
            return "DEFAULT CSV PATH SET TO: " + path;
        }

        [ExcelFunction(Name = Constant.FunctionPrefix + "DataTable.Put", IsVolatile = true)]
        public static string fUtilityPutDataTable(string handle, object[,] data)
        {
            ObjectMapInterface.AddToMap(handle, ExcelFriendlyConversion.ConvertObjectArrayToDataTable(handle, data));
            return handle;
        }

        [ExcelFunction(Name = Constant.FunctionPrefix + "DataTable.Query", IsVolatile = true)]
        public static object[,] fUtilityQueryDataTable(string sqlStringOrHandle, object noHeadersInput, object csvPathOptional)
        {
            string sql = ObjectInfo.DetermineIfSqlHandleOrSqlQuery(sqlStringOrHandle);
            var csvPath = Optional.GetDefaultCsvPath(csvPathOptional);

            object[,] output;
            bool noHeaders = Optional.Check(noHeadersInput, false);
            var tableHandle = ParsingFunctionality.GetTableNameFromSqlQuery(sql);

            if (ObjectMap.Map.Keys.Contains(tableHandle) == false)
            {
                string errorMessage = "ERROR: DataTable + " + tableHandle + " does not exist in the object map.";
                output = new object[1, 1];
                output[0, 0] = errorMessage;
                return output;
            }
            else
            {
                var table = ObjectMapInterface.GetFromMap<DataTable>(tableHandle);
                var newTable = QueryCSV.QueryDataTable(table, sql, csvPath);
                output = ExcelFriendlyConversion.ConvertDataTableToObjectArray(newTable, noHeaders);
                return output;
            }
        }

        [ExcelFunction(Name = Constant.FunctionPrefix + "DataTable.QueryAndStore", IsVolatile = true)]
        public static string fUtilityQueryDataTable(string handle, string sqlStringOrHandle, object csvPathOptional)
        {
            string sql = ObjectInfo.DetermineIfSqlHandleOrSqlQuery(sqlStringOrHandle);
            var tableHandle = ParsingFunctionality.GetTableNameFromSqlQuery(sql);
            var csvPath = Optional.GetDefaultCsvPath(csvPathOptional);

            if (ObjectMap.Map.Keys.Contains(tableHandle) == false)
                return "ERROR: DataTable + " + tableHandle + " does not exist in the object map.";
            else
            {
                var table = ObjectMapInterface.GetFromMap<DataTable>(tableHandle);
                var newTable = QueryCSV.QueryDataTable(table, sql, csvPath);
                newTable.TableName = handle;
                ObjectMapInterface.AddToMap(handle, newTable);
                return handle;
            }
        }

        [ExcelFunction(Name = Constant.FunctionPrefix + "DataTable.QueryMultiple", IsVolatile = true)]
        public static object[,] fUtilityQueryDataTables(string handle, string sqlStringOrHandle, object[] tableHandles, object noHeadersInput, object csvPathOptional)
        {
            List<DataTable> tables = new List<DataTable>();

            var csvPath = Optional.GetDefaultCsvPath(csvPathOptional);
            string sql = ObjectInfo.DetermineIfSqlHandleOrSqlQuery(sqlStringOrHandle);

            bool noHeaders = Optional.Check(noHeadersInput, false);

            for (int i = 0; i < tableHandles.Length; i++)
                tables.Add(ObjectMapInterface.GetFromMap<DataTable>(tableHandles[i].ToString()));

            var newTable = QueryCSV.QueryDataTable(tables, sql, csvPath);
            newTable.TableName = handle;
            var output = ExcelFriendlyConversion.ConvertDataTableToObjectArray(newTable, false);
            return output;
        }

        [ExcelFunction(Name = Constant.FunctionPrefix + "DataTable.QueryMultipleAndStore", IsVolatile = true)]
        public static string fUtilityQueryDataTablesAndStore(string handle, string sqlStringOrHandle, object[] tableHandles, object csvPathOptional)
        {
            List<DataTable> tables = new List<DataTable>();

            var csvPath = Optional.GetDefaultCsvPath(csvPathOptional);
            string sql = ObjectInfo.DetermineIfSqlHandleOrSqlQuery(sqlStringOrHandle);

            for (int i = 0; i < tableHandles.Length; i++)
                tables.Add(ObjectMapInterface.GetFromMap<DataTable>(tableHandles[i].ToString()));

            var newTable = QueryCSV.QueryDataTable(tables, sql, csvPath);
            newTable.TableName = handle;
            ObjectMapInterface.AddToMap(handle, newTable);
            return handle;
        }

        [ExcelFunction(Name = Constant.FunctionPrefix + "DataTable.Get", IsVolatile = true)]
        public static object[,] fUtilityGetDataTable(string handle, object noHeadersInput)
        {

            bool noHeaders = Optional.Check(noHeadersInput, false);

            if (ExcelDnaUtil.IsInFunctionWizard())
                return new object[0, 0];
            else
            {
                var output = ObjectMapInterface.GetFromMap<DataTable>(handle);
                return ExcelFriendlyConversion.ConvertDataTableToObjectArray(output,noHeaders);
            }
        }

        [ExcelFunction(Name = Constant.FunctionPrefix + "DataTable.PutFromSQL", IsVolatile = true)]
        public static string fUtilityDataBasePutDataTableFromSQL(string handle, string connectionString, string sqlStringOrHandle)
        {
            string sql = ObjectInfo.DetermineIfSqlHandleOrSqlQuery(sqlStringOrHandle);

            if (ExcelDnaUtil.IsInFunctionWizard())
                return "no wizard.";
                var table = DatabaseFunctionality.GetDataTableFromSql(sql, connectionString);
            ObjectMapInterface.AddToMap(handle, table);
            table.TableName = handle;
            return handle;
        }

        [ExcelFunction(Name = Constant.FunctionPrefix + "DataTable.PutFromSQLOleDb", IsVolatile = true)]
        public static string fUtilityDataBasePutDataTableFromSQLOleDb(string handle, string connectionString, string sqlStringOrHandle)
        {
            string sql = ObjectInfo.DetermineIfSqlHandleOrSqlQuery(sqlStringOrHandle);

            if (ExcelDnaUtil.IsInFunctionWizard())
                return "no wizard.";
            var table = DatabaseFunctionality.GetDataTableFromSQLOleDb(sql, connectionString);
            table.TableName = handle;
            ObjectMapInterface.AddToMap(handle, table);
            return handle;
        }

        [ExcelFunction(Name = Constant.FunctionPrefix + "DataTable.PutFromSQLADODB", IsVolatile = true)]
        public static string fUtilityDataBasePutDataTableFromSQLADODB(string handle, string connectionString, string sqlStringOrHandle)
        {
            string sql = ObjectInfo.DetermineIfSqlHandleOrSqlQuery(sqlStringOrHandle);

            if (ExcelDnaUtil.IsInFunctionWizard())
                return "no wizard.";
            var table = DatabaseFunctionality.GetDataTableFromSQLAdoDb(sql, connectionString);
            table.TableName = handle;
            ObjectMapInterface.AddToMap(handle, table);
            return handle;
        }
    }

    public class ObjectInfo
    {
        [ExcelFunction(Name = Constant.FunctionPrefix + "Objects.ListMemoryMap", IsVolatile = true)]
        public static object[,] fUtilityListMemoryMap()
        {
            var keyCount = ObjectMap.Map.Keys.Count;
            object[,] output = new object[keyCount+1, 2];

            output[0, 0] = "HANDLE";
            output[0, 1] = "TYPE";

            var i = 1;

            foreach (string key in ObjectMap.Map.Keys)
            {
                output[i, 0] = key;
                output[i, 1] = ObjectMap.Map[key].GetType().Name;
                i = i + 1;
            }

            return output;
        }

        public static string DetermineIfSqlHandleOrSqlQuery(string sqlHandleOrString)
        {
            // Should be modified when more clever object store is implemented

            if (ObjectMap.Map.ContainsKey(sqlHandleOrString))
                return ObjectMapInterface.GetFromMap<string>(sqlHandleOrString);
            else
                return sqlHandleOrString;
        }
    }

    public class DataUtility
    {
        [ExcelFunction(Name = Constant.FunctionPrefix + "Table.SortByHeader")]
        public static object[,] fUtilityTableSortByHeader(object[,] table, object[] headers, object SortAscending, object noHeadersInput)
        {
            bool ascending = Optional.Check(SortAscending, true);
            bool noHeaders = Optional.Check(noHeadersInput, false);

            var dataTable = ExcelFriendlyConversion.ConvertObjectArrayToDataTable("TEMPTABLE", table);
            var header = "";

            for (int i = 0; i < headers.Length-1; i ++)
                header = headers[i].ToString() + ", " + header;

            header = header + headers[headers.Length - 1];

            var direction = (ascending) ? "ASC" : "DESC";

            dataTable.DefaultView.Sort = header + " " + direction;

            return ExcelFriendlyConversion.ConvertDataTableToObjectArray(dataTable.DefaultView.ToTable());
        }

        [ExcelFunction(Name = Constant.FunctionPrefix + "DataTable.SortAndGet")]
        public static object[,] fUtilityDataTableSortAndGet(string handle, object[] headers, object SortAscending, object noHeadersInput)
        {
            var table = ObjectMapInterface.GetFromMap<DataTable>(handle);
            var output = ExcelFriendlyConversion.ConvertDataTableToObjectArray(table);
            return fUtilityTableSortByHeader(output, headers, SortAscending, noHeadersInput);
        }


        [ExcelFunction(Name = Constant.FunctionPrefix + "Array.GetDistinctValue", IsVolatile = true)]
        public static object[,] fUtilityGetDistinctValues(object[] array)
        {
            var list = array.ToList<object>();

            return ExcelFriendlyConversion.ArrayToVerticalObject(list.Distinct().ToArray());
        }

        public static object[] removeEmptyEntries(object[] array)
        {
            List<object> output = new List<object>();

            foreach (object entry in array)
            {
                if (entry is ExcelDna.Integration.ExcelEmpty == false)
                    output.Add(entry);
            }

            return output.ToArray();
        }

        [ExcelFunction(Name = Constant.FunctionPrefix + "Array.RemoveEmptyEntries", IsVolatile = true)]
        public static object[,] fUtilityRemoveEmptyEntries(object[] array)
         {
                List<object> output = new List<object>();

                foreach (object entry in array)
                {
                    if (entry is ExcelDna.Integration.ExcelEmpty == false)
                        output.Add(entry);
                }
                            
            return ExcelFriendlyConversion.ArrayToVerticalObject(output.ToArray());
        }

        [ExcelFunction(Name = Constant.FunctionPrefix + "Array.IgnoreNA", IsVolatile = true)]
        public static object[,] fUtilityArrayIgnoreNA(object[] array)
        {
            List<object> output = new List<object>();

            foreach (object entry in array)
            {
                if (entry is ExcelDna.Integration.ExcelError == false)
                    output.Add(entry);
            }

            return ExcelFriendlyConversion.ArrayToVerticalObject(output.ToArray());
        }

        [ExcelFunction(Name = Constant.FunctionPrefix + "Array.GetIntersection", IsVolatile = true)]
        public static object[,] fUtilityGetIntersection(object[] array1, object[] array2)
        {
            var list1 = array1.ToList();
            var list2 = array2.ToList();

            var group1 =
                from n in list1
                group n by n
                into g
                select new { g.Key, count = g.Count() };

            var group2 =
                from n in list2
                group n by n
                into g
                select new { g.Key, count = g.Count() };

            var joined =
                from b in group2
                join a in group1 on b.Key equals a.Key
                select new { b.Key, Count = Math.Min(b.count, a.count) };

            return ExcelFriendlyConversion.ArrayToVerticalObject(joined.SelectMany(a => Enumerable.Repeat(a.Key, a.Count)).ToArray());

        }
    }

    internal static class Optional
    {
        internal static bool Check(object arg, bool defaultValue)
        {
            if (arg is bool)
                return (bool)arg;
            else if (arg is ExcelMissing)
                return defaultValue;
            else
                throw new ArgumentException();

            // Perhaps check for other types and do whatever you think is right ....
            //else if (arg is double)
            //    return "Double: " + (double)arg;
            //else if (arg is bool)
            //    return "Boolean: " + (bool)arg;
            //else if (arg is ExcelError)
            //    return "ExcelError: " + arg.ToString();
            //else if (arg is object[,])
            //    // The object array returned here may contain a mixture of types,
            //    // reflecting the different cell contents.
            //    return string.Format("Array[{0},{1}]", 
            //      ((object[,])arg).GetLength(0), ((object[,])arg).GetLength(1));
            //else if (arg is ExcelEmpty)
            //    return "<<Empty>>"; // Would have been null
            //else if (arg is ExcelReference)
            //  // Calling xlfRefText here requires IsMacroType=true for this function.
            //				return "Reference: " + 
            //                     XlCall.Excel(XlCall.xlfReftext, arg, true);
            //			else
            //				return "!? Unheard Of ?!";
        }
        internal static string GetDefaultCsvPath(object arg)
        {
            if (arg is ExcelMissing)
                return ObjectMapInterface.GetFromMap<string>(Constant.DefaultCsvHandle);
            else
                return arg.ToString();
        }

    }
}
