using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

namespace fUtility
{
    public static class ObjectMap
    {
        public static IDictionary<string, string> MapConnectionStrings = new Dictionary<string, string>();

        public static IDictionary<string, object> Map = new Dictionary<string, object>();
    }
}
