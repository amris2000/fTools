using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

namespace fUtility
{
    public static class ExcelFriendlyConversion
    {
        public static object[,] ArrayToVerticalObject(object[] array)
        {
            object[,] output = new object[array.Length, 1];

            for (int i = 0; i<array.Length; i++)
                output[i, 0] = array[i];

            return output;
        }

        public static object[,] ArrayToVerticalObject(List<object> array)
        {
            return ArrayToVerticalObject(array.ToArray());
        }

        public static object[,] ConvertDataTableToObjectArray(DataTable table)
        {
            var rows = table.Rows.Count;
            var cols = table.Columns.Count;
            var output = new object[rows + 1, cols];

            // Fields
            for (int i = 0; i < cols; i++)
                output[0, i] = table.Columns[i].ColumnName;

            // Values
            for (int i = 0; i < cols; i ++)
            {
                for (int j = 1; j < rows+1; j++)
                {
                    if (table.Rows[j - 1][i] == DBNull.Value)
                        output[j, i] = "";
                    else
                        output[j, i] = table.Rows[j-1][i];
                }
            }

            return output;
        }

        public static DataTable ConvertObjectArrayToDataTable(string name, object[,] array)
        {
            var table = new DataTable(name);
            var rows = array.GetLength(0);
            var cols = array.GetLength(1);


            var headers = new object[cols];

            for (int i = 0; i < cols; i++)
            {
                var column = new DataColumn();
                column.DataType = System.Type.GetType(typeof(object).ToString());
                column.ColumnName = array[0, i].ToString().ToUpper();
                table.Columns.Add(column);
            }

            for (int i = 1; i<rows; i++)
            {
                var dataRow = table.NewRow();
                for (int j = 0; j<cols; j++)
                    dataRow[array[0,j].ToString().ToUpper()] = array[i,j];
                table.Rows.Add(dataRow);
            }

            return table;
        }

    }
}
