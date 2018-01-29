using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using System.Data;
using fUtility;
using System.Globalization;

namespace fUtilityExcel
{
    static class DEFINITIONS
    {
        public const string DATATABLE_FIELD = "DATATABLE";
        public const string SQL_FIELD = "SQL";
    }

    public class Finance
    {
        [ExcelFunction(Name = Constant.FunctionPrefix + "Finance.GetFx", IsVolatile = true)]
        public static double FinanceGetFx(string baseCcy, string varCcy)
        {
            if (ExcelDnaUtil.IsInFunctionWizard())
            {
                return 0.0;
            }
            else
            {
                var fx = fUtility.CoinMarketCapApi.GetJson.GetFxRate(baseCcy, varCcy);
                return fx.FxRate;
            }

        }

        [ExcelFunction(Name = Constant.FunctionPrefix + "Finance.GetCrypto", IsVolatile = true)]
        public static double FinanceGetCrypto(string crypto, object currencyInput)
        {
            if (ExcelDnaUtil.IsInFunctionWizard())
                return 0.0;
            else
            {
                string currency = CheckCurrencyInput(currencyInput);
                return fUtility.CoinMarketCapApi.GetJson.GetCryptoRate(crypto, currency);
            }
        }

        public static string CheckCurrencyInput(object currencyInput)
        {
            bool noCurrency = Optional.Check(currencyInput, true);
            string currency;

            if (noCurrency)
                currency = "EUR";
            else
                currency = currencyInput.ToString();

            return currency;
        }

        [ExcelFunction(Name = Constant.FunctionPrefix + "Finance.GetCryptoCross", IsVolatile = true)]
        public static double FinanceGetCryptoCross(string cryptoBase, string cryptoVariable)
        {
            return fUtility.CoinMarketCapApi.GetJson.GetCryptoCross(cryptoBase, cryptoVariable);
        }




        [ExcelFunction(Name = Constant.FunctionPrefix + "Finance.GetCryptoInfo", IsVolatile = true)]
        public static object[] FinanceGetCryptoInfo(string crypto, object currencyInput)
        {
            if (ExcelDnaUtil.IsInFunctionWizard())
                return new object[1] { 0.0 };
            else
            {
                string currency = CheckCurrencyInput(currencyInput);
                var coin = fUtility.CoinMarketCapApi.GetJson.GetCoinInfoFromSource(crypto, currency);
                List<object> results = new List<object>();


                double price, change1h, change24h, change7d, supply, rank;

                price = fUtility.CoinMarketCapApi.GetJson.TryConvertDouble(coin[0].price_eur);
                double.TryParse(coin[0].percent_change_1h, NumberStyles.Any, CultureInfo.InvariantCulture, out change1h);
                double.TryParse(coin[0].percent_change_24h, NumberStyles.Any, CultureInfo.InvariantCulture, out change24h);
                double.TryParse(coin[0].percent_change_7d, NumberStyles.Any, CultureInfo.InvariantCulture, out change7d);
                double.TryParse(coin[0].available_supply, NumberStyles.Any, CultureInfo.InvariantCulture, out supply);
                double.TryParse(coin[0].rank, NumberStyles.Any, CultureInfo.InvariantCulture, out rank);

                results.Add(coin[0].symbol);
                results.Add(price);
                results.Add(change1h);
                results.Add(change24h);
                results.Add(change7d);
                results.Add(supply);
                results.Add(rank);

                return results.ToArray();
            }
        }

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
