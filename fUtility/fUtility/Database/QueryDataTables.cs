using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data;

namespace fUtility
{
    public static class QueryDataTables
    {
        public static DataTable NewTableFromQuery(string query, DataTable table)
        {
            return table.Select(query).CopyToDataTable();
        }

        public static DataTable NewTableFromComputeOnTable(string expression, string filter, DataTable table)
        {
            return (DataTable)table.Compute(expression, filter);
        }
    }
}
