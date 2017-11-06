using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace fUtility
{
    public static class DictionaryExtensions
    {
        public static T Get<T>(this IDictionary<string, object> instance, string name)
        {
            return (T)instance[name];
        }
    }
}
