using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Globalization;
using System.Text.RegularExpressions;

namespace fUtility.Database
{
    // Example of usage

    //// set culture to invariant
    //StringConverter.Default.Culture = CultureInfo.InvariantCulture;
    //// add custom converter to default, it will match strings starting with CUSTOM: and return MyCustomClass
    // StringConverter.Default.AddConverter(c => c.StartsWith("CUSTOM:"), c => new MyCustomClass(c));
    //var items = new[] { "1", "4343434343", "3.33", "true", "false", "2014-10-10 22:00:00", "CUSTOM: something" };
    //foreach (var item in items) {
    //    object result;
    //    if (StringConverter.Default.TryConvert(item, out result)) {
    //        Console.WriteLine(result);
    //    }
    //}

    public class StringConverter
    {
        // delegate for TryParse(string, out T)
        public delegate bool TypedConvertDelegate<T>(string value, out T result);
        // delegate for TryParse(string, out object)
        private delegate bool UntypedConvertDelegate(string value, out object result);
        private readonly List<UntypedConvertDelegate> _converters = new List<UntypedConvertDelegate>();
        // default converter, lazyly initialized
        private static readonly Lazy<StringConverter> _default = new Lazy<StringConverter>(CreateDefault, true);

        public static StringConverter Default => _default.Value;

        private static StringConverter CreateDefault()
        {
            var d = new StringConverter();
            // add reasonable default converters for common .NET types. Don't forget to take culture into account, that's
            // important when parsing numbers\dates.
            d.AddConverter<bool>(bool.TryParse);
            d.AddConverter((string value, out byte result) => byte.TryParse(value, NumberStyles.Integer, d.Culture, out result));
            d.AddConverter((string value, out short result) => short.TryParse(value, NumberStyles.Integer, d.Culture, out result));
            d.AddConverter((string value, out int result) => int.TryParse(value, NumberStyles.Integer, d.Culture, out result));
            d.AddConverter((string value, out long result) => long.TryParse(value, NumberStyles.Integer, d.Culture, out result));
            d.AddConverter((string value, out float result) => float.TryParse(value, NumberStyles.Number, d.Culture, out result));
            d.AddConverter((string value, out double result) => double.TryParse(value, NumberStyles.Number, d.Culture, out result));
            d.AddConverter((string value, out DateTime result) => DateTime.TryParse(value, d.Culture, DateTimeStyles.None, out result));
            return d;
        }

        //
        public CultureInfo Culture { get; set; } = CultureInfo.CurrentCulture;

        public void AddConverter<T>(Predicate<string> match, Func<string, T> converter)
        {
            // create converter from match predicate and convert function
            _converters.Add((string value, out object result) => {
                if (match(value))
                {
                    result = converter(value);
                    return true;
                }
                result = null;
                return false;
            });
        }

        public void AddConverter<T>(Regex match, Func<string, T> converter)
        {
            // create converter from match regex and convert function
            _converters.Add((string value, out object result) => {
                if (match.IsMatch(value))
                {
                    result = converter(value);
                    return true;
                }
                result = null;
                return false;
            });
        }

        public void AddConverter<T>(TypedConvertDelegate<T> constructor)
        {
            // create converter from typed TryParse(string, out T) function
            _converters.Add(FromTryPattern<T>(constructor));
        }

        public bool TryConvert(string value, out object result)
        {
            if (this != Default)
            {
                // if this is not a default converter - first try convert with default
                if (Default.TryConvert(value, out result))
                    return true;
            }
            // then use local converters. Any will return after the first match
            object tmp = null;
            bool anyMatch = _converters.Any(c => c(value, out tmp));
            result = tmp;
            return anyMatch;
        }

        private static UntypedConvertDelegate FromTryPattern<T>(TypedConvertDelegate<T> inner)
        {
            return (string value, out object result) => {
                T tmp;
                if (inner.Invoke(value, out tmp))
                {
                    result = tmp;
                    return true;
                }
                else
                {
                    result = null;
                    return false;
                }
            };
        }
    }
}
