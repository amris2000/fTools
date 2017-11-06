using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

namespace fUtility
{
    public static class DataTableUtility
    {
        public static List<string> GetFields(DataTable table)
        {
            var fields = new List<string>();
            var fieldCount = table.Columns.Count;

            for (int i = 0; i < fieldCount; i++)
                fields.Add(table.Columns[i].ColumnName);

            return fields;
        }
    }
}
