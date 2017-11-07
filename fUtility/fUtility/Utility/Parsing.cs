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
            // This needs to be improved.. 

            var queryArray = query.Split(' ');
            string result = null;

            for (int i = 0; i < queryArray.Length; i++)
            {
                if (queryArray[i].ToUpper() == "FROM")
                {
                    result = queryArray[i + 1];
                    break;
                }
            }

            if (queryArray == null)
                throw new InvalidOperationException("Could not find a tablename in query");

            return result;

            //string tableName = "";
            //query = query.ToUpper();
            //int positionIndex = query.IndexOf("FROM");

            //tableName = query.Substring(positionIndex);
            //tableName = tableName.Replace("FROM ", string.Empty);

            //if (tableName.Contains(" "))
            //    tableName = tableName.Left(tableName.IndexOf(" "));
            
            //tableName = tableName.Trim();   

            //return tableName;
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
