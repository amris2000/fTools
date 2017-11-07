using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Threading.Tasks;
using ADODB;

namespace fUtility
{
    public static class DataTableTest
    {
        public static DataTable TestTable()
        {

            var dataTable = new DataTable("MYTABLE");
            Random rnd = new Random();

            var rows = 5;
            var cols = 3;


            for (int i = 0; i<cols; i++)
            {
                var column = new DataColumn();
                column.DataType = System.Type.GetType(typeof(object).ToString());
                column.ColumnName = "COL" + i;
                dataTable.Columns.Add(column);
            }

            for (int i = 0; i<rows; i++)
            {
                var dataRow = dataTable.NewRow();
                for (int j = 0; j<cols; j++)
                    dataRow["COL" + j] = rnd.Next(0, 10);
                dataTable.Rows.Add(dataRow);
            }

            return dataTable;            

        }
    }
}
