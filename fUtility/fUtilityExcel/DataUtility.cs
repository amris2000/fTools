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
    public static class Constant
    {
        public const string FunctionPrefix = "fUtility.";
    }

    public class DataTableFunctions
    {

        [ExcelFunction(Name = Constant.FunctionPrefix + "DataTables.GetFields", IsVolatile = true)]
        public static object[,] fUtilityDataTablesGetFields(string handle)
        {
            var table = ObjectMapInterface.GetFromMap<DataTable>(handle);
            var fields = DataTableUtility.GetFields(table);
            var output = new object[fields.Count, 1];
            for (int i = 0; i < fields.Count; i++)
                output[i, 0] = fields[i];

            return output;
        }

        [ExcelFunction(Name = Constant.FunctionPrefix + "DataTables.GetColumnCount", IsVolatile = true)]
        public static int fUtilityDataTablesGetColumnCount(string handle)
        {
            var table = ObjectMapInterface.GetFromMap<DataTable>(handle);
            return table.Columns.Count;
        }

        [ExcelFunction(Name = Constant.FunctionPrefix + "DataTables.GetRowCount", IsVolatile = true)]
        public static int fUtilityDataTablesGetRowCount(string handle)
        {
            var table = ObjectMapInterface.GetFromMap<DataTable>(handle);
            return table.Rows.Count;
        }

        [ExcelFunction(Name = Constant.FunctionPrefix + "ReadSql", IsVolatile = true)]
        public static string fUtilityReadSql(object[] array)
        {
            return ParsingFunctionality.ReadSqlFromArray(array);
        }

        [ExcelFunction(Name = Constant.FunctionPrefix + "PutDataTable", IsVolatile = true)]
        public static string fUtilityPutDataTable(string handle, object[,] data)
        {
            ObjectMapInterface.AddToMap(handle, ExcelFriendlyConversion.ConvertObjectArrayToDataTable(handle, data));
            return handle;
        }

        [ExcelFunction(Name = Constant.FunctionPrefix + "QueryDataTable", IsVolatile = true)]
        public static object[,] fUtilityQueryDataTable(string sql, string csvPath)
        {
            object[,] output;
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
                output = ExcelFriendlyConversion.ConvertDataTableToObjectArray(newTable);
                return output;
            }
        }

        [ExcelFunction(Name = Constant.FunctionPrefix + "QueryDataTableAndStore", IsVolatile = true)]
        public static string fUtilityQueryDataTable(string handle, string sql, string csvPath)
        {
            var tableHandle = ParsingFunctionality.GetTableNameFromSqlQuery(sql);

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

        [ExcelFunction(Name = Constant.FunctionPrefix + "QueryDataTables", IsVolatile = true)]
        public static object[,] fUtilityQueryDataTables(string handle, string sql, string csvPath, object[] tableHandles)
        {
            List<DataTable> tables = new List<DataTable>();

            for (int i = 0; i < tableHandles.Length; i++)
                tables.Add(ObjectMapInterface.GetFromMap<DataTable>(tableHandles[i].ToString()));

            var newTable = QueryCSV.QueryDataTable(tables, sql, csvPath);
            newTable.TableName = handle;
            var output = ExcelFriendlyConversion.ConvertDataTableToObjectArray(newTable);
            return output;
        }

        [ExcelFunction(Name = Constant.FunctionPrefix + "QueryDataTablesAndStore", IsVolatile = true)]
        public static string fUtilityQueryDataTablesAndStore(string handle, string sql, string csvPath, object[] tableHandles)
        {
            List<DataTable> tables = new List<DataTable>();

            for (int i = 0; i < tableHandles.Length; i++)
                tables.Add(ObjectMapInterface.GetFromMap<DataTable>(tableHandles[i].ToString()));

            var newTable = QueryCSV.QueryDataTable(tables, sql, csvPath);
            ObjectMapInterface.AddToMap(handle, newTable);
            return handle;
        }


        [ExcelFunction(Name = Constant.FunctionPrefix + "GetDataTable", IsVolatile = true)]
        public static object[,] fUtilityGetDataTable(string handle)
        {
            if (ExcelDnaUtil.IsInFunctionWizard())
                return new object[0, 0];
            else
            {
                var output = ObjectMapInterface.GetFromMap<DataTable>(handle);
                return ExcelFriendlyConversion.ConvertDataTableToObjectArray(output);
            }
        }

        [ExcelFunction(Name = Constant.FunctionPrefix + "Database.PutDataTableFromSQL", IsVolatile = true)]
        public static string fUtilityDataBasePutDataTableFromSQL(string handle, string connectionString, string sql)
        {
            if (ExcelDnaUtil.IsInFunctionWizard())
                return "no wizard.";
                var table = DatabaseFunctionality.GetDataTableFromSql(sql, connectionString);
            ObjectMapInterface.AddToMap(handle, table);
            return handle;
        }
    }

    public class ObjectInfo
    {
        [ExcelFunction(Name = Constant.FunctionPrefix + "ListMemoryMap", IsVolatile = true)]
        public static object[,] fUtilityListMemoryMap()
        {
            var keyCount = ObjectMap.Map.Keys.Count;
            object[,] output = new object[keyCount+1, 3];

            output[0, 0] = "KEY #";
            output[0, 1] = "HANDLE";
            output[0, 2] = "TYPE";

            var i = 1;

            foreach (string key in ObjectMap.Map.Keys)
            {
                output[i, 0] = i;
                output[i, 1] = key;
                output[i, 2] = ObjectMap.Map[key].GetType().Name;
                i = i + 1;
            }

            return output;
        }
    }


    public class DataUtility
    {
        [ExcelFunction(Name = Constant.FunctionPrefix + "GetDistinctValue", IsVolatile = true)]
        public static object[,] fUtilityGetDistinctValues(object[] array)
        {
            var list = array.ToList<object>();

            return ExcelFriendlyConversion.ArrayToVerticalObject(list.Distinct().ToArray());
        }

        [ExcelFunction(Name = Constant.FunctionPrefix + "RemoveEmptyEntries", IsVolatile = true)]
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

        [ExcelFunction(Name = Constant.FunctionPrefix + "GetIntersection", IsVolatile = true)]
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
}
