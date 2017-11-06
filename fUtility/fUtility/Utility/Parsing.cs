using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace fUtility
{
    public static class ParsingFunctionality
    {
        public static string GetTableNameFromSqlQuery(string query)
        {
            string tableName = "";
            query = query.ToUpper();
            int positionIndex = query.IndexOf("FROM");

            tableName = query.Substring(positionIndex);
            tableName = tableName.Replace("FROM ", string.Empty);
            tableName = tableName.Left(tableName.IndexOf(" "));
            tableName = tableName.Trim();   

            return tableName;
        }

        public static string ReadSqlFromArray(object[] query)
        {
            string result = "";

            for (int i = 0; i < query.Length; i++)
                result = result + " " + query[i].ToString();

            return result;
        }
    }

    
}
