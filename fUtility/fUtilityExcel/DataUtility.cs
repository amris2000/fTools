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
        public const string FunctionPrefix = "ft.";
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
            return ParsingFunctionality.ReadSqlFromArray(array);
        }

        [ExcelFunction(Name = Constant.FunctionPrefix + "DataTable.Put", IsVolatile = true)]
        public static string fUtilityPutDataTable(string handle, object[,] data)
        {
            ObjectMapInterface.AddToMap(handle, ExcelFriendlyConversion.ConvertObjectArrayToDataTable(handle, data));
            return handle;
        }

        [ExcelFunction(Name = Constant.FunctionPrefix + "DataTable.Query", IsVolatile = true)]
        public static object[,] fUtilityQueryDataTable(string sql, string csvPath, object noHeadersInput)
        {
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

        [ExcelFunction(Name = Constant.FunctionPrefix + "DataTable.QueryMultiple", IsVolatile = true)]
        public static object[,] fUtilityQueryDataTables(string handle, string sql, string csvPath, object[] tableHandles, object noHeadersInput)
        {
            List<DataTable> tables = new List<DataTable>();

            bool noHeaders = Optional.Check(noHeadersInput, false);

            for (int i = 0; i < tableHandles.Length; i++)
                tables.Add(ObjectMapInterface.GetFromMap<DataTable>(tableHandles[i].ToString()));

            var newTable = QueryCSV.QueryDataTable(tables, sql, csvPath);
            newTable.TableName = handle;
            var output = ExcelFriendlyConversion.ConvertDataTableToObjectArray(newTable, false);
            return output;
        }

        [ExcelFunction(Name = Constant.FunctionPrefix + "DataTable.QueryMultipleAndStore", IsVolatile = true)]
        public static string fUtilityQueryDataTablesAndStore(string handle, string sql, string csvPath, object[] tableHandles)
        {
            List<DataTable> tables = new List<DataTable>();

            for (int i = 0; i < tableHandles.Length; i++)
                tables.Add(ObjectMapInterface.GetFromMap<DataTable>(tableHandles[i].ToString()));

            var newTable = QueryCSV.QueryDataTable(tables, sql, csvPath);
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
        [ExcelFunction(Name = Constant.FunctionPrefix + "Objects.ListMemoryMap", IsVolatile = true)]
        public static object[,] fUtilityListMemoryMap()
        {
            var keyCount = ObjectMap.Map.Keys.Count;
            object[,] output = new object[keyCount+1, 3];

            output[0, 0] = "#";
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

    }
}
