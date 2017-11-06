using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using fUtility;

namespace fUtilityConsole
{
    class Program
    {

        static void Main(string[] args)
        {
            Console.WriteLine("Start");
            // var dataTable = DataTableTest.TestTable();
            // var output = ExcelFriendlyConversion.ConvertDataTableToObjectArray(dataTable);
            //var dataTable2 = ExcelFriendlyConversion.ConvertObjectArrayToDataTable("name", output);

            string query = "SELECT NAME, INCOME FROM MYTABLENAME WHERE SOMETHING HAPPENS";
            Console.WriteLine(ParsingFunctionality.GetTableNameFromSqlQuery(query));



            Console.WriteLine("DataTables");
            Console.ReadLine();

        }
    }
}
