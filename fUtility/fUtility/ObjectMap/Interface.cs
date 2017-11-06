using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

namespace fUtility
{
    public static class ObjectMapInterface
    {
        public static void AddToMap(string key, object data)
        {
            if (ObjectMap.Map.Keys.Contains(key))
                ObjectMap.Map.Remove(key);

            ObjectMap.Map.Add(key, data);
        }


        public static T GetFromMap<T>(string key)
        {
            if (ObjectMap.Map.Keys.Contains(key) == false)
                throw new InvalidOperationException("MAP DOES NOT CONTAIN KEY " + key + ".");
            else
                return (T)ObjectMap.Map[key];
        }

    }
}
