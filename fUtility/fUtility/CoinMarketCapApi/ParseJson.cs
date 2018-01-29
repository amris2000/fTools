using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.IO;
using System.Web.Script.Serialization;
using System.Globalization;
using Newtonsoft.Json;

namespace fUtility.CoinMarketCapApi
{
    public static class GetJson
    {
        public static double TryConvertDouble(string input)
        {
            double output;
            double.TryParse(input, NumberStyles.Any, CultureInfo.InvariantCulture, out output);
            return output;
        }

        public static double GetCryptoRate(string crypto, string currency)
        {
            var coin = fUtility.CoinMarketCapApi.GetJson.GetCoinInfoFromSource(crypto, currency);
            double fxRate;
            double.TryParse(coin[0].price_eur, NumberStyles.Any, CultureInfo.InvariantCulture, out fxRate);
            return fxRate;
        }

        public static double GetCryptoCross(string cryptoBase, string cryptoVariable)
        {
            var baseCoin = fUtility.CoinMarketCapApi.GetJson.GetCoinInfoFromSource(cryptoBase, "EUR");
            var variableCoin = fUtility.CoinMarketCapApi.GetJson.GetCoinInfoFromSource(cryptoVariable, "EUR");

            var fxBase = fUtility.CoinMarketCapApi.GetJson.TryConvertDouble(baseCoin[0].price_eur);
            var fxVariable = fUtility.CoinMarketCapApi.GetJson.TryConvertDouble(variableCoin[0].price_eur);

            return fxBase / fxVariable;
        }


        public static string GetStringFor(string cryptoCurrency, string currency)
        {
            return "https://api.coinmarketcap.com/v1/ticker/" + cryptoCurrency + "/?convert=" + currency;
        }

        public static string Get(string cryptoCurrency, string currency)
        {
            string address = GetStringFor(cryptoCurrency, currency);
            var json = new WebClient().DownloadString(address);
   
            return json;
        }

        public class FxApi
        {
            public string BaseCcy { get; set; }
            public string VarCcy { get; set; }
            public double FxRate { get; set; }

            public FxApi(string baseCcy, string varCcy, double rate)
            {
                BaseCcy = baseCcy;
                VarCcy = varCcy;
                FxRate = rate;
            }

            public void Print()
            {
                Console.WriteLine("Fx rate for {0}{1}: {2}", BaseCcy, VarCcy, FxRate);
            }
        }

        public class CoinApi
        {
            public string id { get; set; }
            public string name { get; set; }
            public string symbol { get; set; }
            public string rank { get; set; }
            public string price_usd { get; set; }
            [JsonProperty(PropertyName = "24h_volume_usd")]   //since in c# variable names cannot begin with a number, you will need to use an alternate name to deserialize
            public string volume_usd_24h { get; set; }
            public string market_cap_usd { get; set; }
            public string available_supply { get; set; }
            public string total_supply { get; set; }
            public string percent_change_1h { get; set; }
            public string percent_change_24h { get; set; }
            public string percent_change_7d { get; set; }
            public string last_updated { get; set; }
            public string price_eur { get; set; }
        }

        public static FxApi GetFxRate(string baseCurrency, string variableCurrency)
        {
            string apiKey = "4N9HN7NLQGYLJHEB";
            string url = "https://www.alphavantage.co/query?function=CURRENCY_EXCHANGE_RATE&from_currency=" + baseCurrency + "&to_currency=" + variableCurrency + "&apikey=" + apiKey;
            //var client = new WebClient();

            HttpWebRequest myReq = (HttpWebRequest)WebRequest.Create(url);
            WebResponse myResp = myReq.GetResponse();

            StreamReader reader = new StreamReader(myResp.GetResponseStream());

            var serializer = new System.Web.Script.Serialization.JavaScriptSerializer();
            var jObject = Newtonsoft.Json.Linq.JObject.Parse(reader.ReadToEnd());
            var data = jObject["Realtime Currency Exchange Rate"].ToObject<Dictionary<string, string>>();

            var baseCcy = data["1. From_Currency Code"];
            //Console.WriteLine(data["2. From_Currency Name"]);
            var varCcy = data["3. To_Currency Code"];
            //Console.WriteLine(data["4. To_Currency Name"]);
            var rate = data["5. Exchange Rate"];
            double FxRate;

            if (Double.TryParse(rate, NumberStyles.Any, CultureInfo.InvariantCulture, out FxRate))
                return new FxApi(baseCcy, varCcy, FxRate);
            else
                throw new InvalidOperationException("Could not parse Fx rate to double.");
        }

        public static List<CoinApi> GetCoinInfoFromSource(string cryptoCurrency, string currency)
        {
            var url = GetStringFor(cryptoCurrency, currency);
            HttpWebRequest WebReq = (HttpWebRequest)WebRequest.Create(string.Format(url));

            WebReq.Method = "GET";

            HttpWebResponse WebResp = (HttpWebResponse)WebReq.GetResponse();

            //Console.WriteLine(WebResp.StatusCode);
            //Console.WriteLine(WebResp.Server);

            string jsonString;
            using (Stream stream = WebResp.GetResponseStream())   //modified from your code since the using statement disposes the stream automatically when done
            {
                StreamReader reader = new StreamReader(stream, System.Text.Encoding.UTF8);
                jsonString = reader.ReadToEnd();
            }

            List<CoinApi> items = JsonConvert.DeserializeObject<List<CoinApi>>(jsonString);
            return items;

        }
    }
}
