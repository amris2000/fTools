using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

namespace fUtility
{
    public class ObjectStoreEntry
    {
        public string ObjectType { get; private set; }
        public object Obj { get; private set; }
        public string StoreField { get; private set; }
        public string Handle { get; private set; }
        public string FullKey { get; private set; }
        public DateTime CreationTime { get; private set; }

        public ObjectStoreEntry(string handle, object obj, string storeField)
        {
            if (storeField == null)
                storeField = PersistentObjects.DEFAULT_FIELD_NAME;

            ObjectType = obj.GetType().ToString().ToUpper();
            Obj = obj;
            StoreField = storeField.ToUpper();
            Handle = handle.ToUpper();
            FullKey = StoreField + "." + Handle;
            CreationTime = DateTime.Now;
        }

    }

    public static class ObjectMapInterface
    {

    }
}
