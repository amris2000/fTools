using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace fUtility
{
    //https://stackoverflow.com/questions/4478464/c-sharp-switch-on-type?noredirect=1&lq=1

    public class TypeSwitch
    {
        Dictionary<Type, Action<object>> matches = new Dictionary<Type, Action<object>>();
        public TypeSwitch Case<T>(Action<T> action)
        {
            matches.Add(typeof(T), (x) => action((T)x)); return this;
        }

        public void Switch(object x)
        {
            if (matches.ContainsKey(x.GetType()) == false)
                Default(x);

            matches[x.GetType()](x);
        }

        public void Default(object x)
        {
            matches[typeof(System.String)](x);
        }

    }
}
