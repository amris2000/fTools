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
    static class DEFINITIONS
    {
        public const string DATATABLE_FIELD = "DATATABLE";
        public const string SQL_FIELD = "SQL";
    }

    public class ExcelUtility
    {
        [ExcelFunction(Name = Constant.FunctionPrefix + "DataTable.GetInfoCsv", IsVolatile = true)]
        public static object[,] DataTableGetInfoCsv(string handle)
        {
            var table = PersistentObjects.GetFromMap<DataTable>(handle, DEFINITIONS.DATATABLE_FIELD);
            return CreateDataTableInfo(table);
        }

        //[ExcelFunction(Name = Constant.FunctionPrefix + "DataTable.GetInfoDb", IsVolatile = true)]
        //public static object[,] DataTableGetInfoDb(string handle)
        //{
        //    var table = PersistentObjects.MyDatabase.GetDataTable(handle);
        //    return CreateDataTableInfo(table);
        //}

        public static object[,] CreateDataTableInfo(DataTable table)
        {
            int columns = table.Columns.Count;

            object[,] output = new object[columns + 5, 2];

            output[0, 0] = "TABLENAME";
            output[0, 1] = table.TableName;
            output[1, 0] = "COLUMNS";
            output[1, 1] = columns;
            output[2, 0] = "ROWS";
            output[2, 1] = table.Rows.Count;
            output[3, 0] = "";
            output[3, 1] = "";

            output[4, 0] = "COLUMN_NAME";
            output[4, 1] = "COLUMN_TYPE";

            int i = 5;

            foreach (DataColumn col in table.Columns)
            {
                output[i, 0] = col.ColumnName;
                output[i, 1] = col.DataType.ToString();
                i = i + 1;
            }

            return output;
        }
    }
}
