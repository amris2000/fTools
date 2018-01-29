using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using fUtility;
using fUtility.CoinMarketCapApi;
using System.Data;

namespace fUtilityConsole
{
    class Program
    {
        static DataTable CreateTestTable()
        {
            var table = new DataTable();
            table.Columns.Add("NAME", typeof(System.String));
            table.Columns.Add("INTVAL", typeof(System.Int32));
            table.Columns.Add("DOUBLEVAL", typeof(System.Double));

            table.Rows.Add("FREDERÆIK", 123, 3.213);
            table.Rows.Add("ANGELAååA", 146, 4312.12);
            table.Rows.Add("HELLE", 999, 1.2353244);

            return table;
        }

        static object[,] CreateTestArray()
        {
            var obj = new Object[4, 3];

            obj[0, 0] = "NAME";
            obj[0, 1] = "INTVAL";
            obj[0, 2] = "DOUBLEVAL";

            obj[1, 0] = "FREDERIK";
            obj[1, 1] = 123;
            obj[1, 2] = 3.213;

            obj[2, 0] = "ANGELA";
            obj[2, 1] = 234;
            obj[2, 2] = 123213.223413;

            obj[3, 0] = "HELLE";
            obj[3, 1] = 9993;
            obj[3, 2] = 4.9913;

            return obj;
        }

        static void PrintTableInfoToConsole(DataTable table)
        {
            Console.WriteLine("--- Table info on table {0}", table.TableName);
            foreach (DataColumn col in table.Columns)
                Console.WriteLine("Datatype of column {0} is {1}", col.ColumnName, col.DataType.ToString());

        }
        

        static void TestSQLite()
        {
            var table = CreateTestTable();
            SQLiteDatabase db = new SQLiteDatabase("MYDB");

            db.CreateTable("MYTABLE", table);
            var table2 = db.InspectDatabase();
            var table3 = db.GetDataTable("MYTABLE");

            var obj = CreateTestArray();
            var table4 = ExcelFriendlyConversion.ConvertObjectArrayToDataTable("MYTABLE4", obj);

            PrintTableInfoToConsole(table4);

            db.CreateTable("MYTABLE4", table4);
            var table5 = db.GetDataTable("MYTABLE4");

            PrintTableInfoToConsole(table5);


            Console.WriteLine("Database created ..");
        }

        static void TestCoinAndFxApi()
        {
            var fx = GetJson.GetFxRate("EUR", "USD");
            fx.Print();

            var results = fUtility.CoinMarketCapApi.GetJson.GetCoinInfoFromSource("ethereum", "EUR");
            Console.WriteLine("{0} {1} {2}", results[0].symbol, results[0].name, results[0].price_eur);
            results = fUtility.CoinMarketCapApi.GetJson.GetCoinInfoFromSource("litecoin", "EUR");
            Console.WriteLine("{0} {1} {2}", results[0].symbol, results[0].name, results[0].price_eur);

            results = fUtility.CoinMarketCapApi.GetJson.GetCoinInfoFromSource("bitcoin", "EUR");
            Console.WriteLine("{0} {1} {2}", results[0].symbol, results[0].name, results[0].price_eur);

            results = fUtility.CoinMarketCapApi.GetJson.GetCoinInfoFromSource("ripple", "EUR");
            Console.WriteLine("{0} {1} {2}", results[0].symbol, results[0].name, results[0].price_eur);

            Console.WriteLine("");

            Console.WriteLine("Cross Rates");
            Console.WriteLine("{0}/{1}: {2}", "ripple", "ethereum", fUtility.CoinMarketCapApi.GetJson.GetCryptoCross("ripple", "ethereum"));
        }


        static void Main(string[] args)
        {
            TestCoinAndFxApi();

            Console.ReadLine();

        }
    }
}
