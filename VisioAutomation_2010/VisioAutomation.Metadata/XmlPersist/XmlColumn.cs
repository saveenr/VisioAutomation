using System.Collections.Generic;
using System.Linq;

namespace XmlPersist
{
    public class XmlColumn
    {
        public string Name;
        public int Ordinal;
        public System.Reflection.PropertyInfo PropertyInfo;

        public string GetStringValue<T>(T o)
        {
            object v = this.PropertyInfo.GetValue(o, null);
            string vs = (string)v;
            return vs;
        }

        public void SetStringValue<T>(T o, string val)
        {
            this.PropertyInfo.SetValue(o, val, null);
        }

        public XmlColumn(System.Reflection.PropertyInfo propinfo)
        {
            this.Name = propinfo.Name;
            this.PropertyInfo = propinfo;

            if (propinfo.PropertyType != typeof(string))
            {
                string msg = string.Format("Property \"{0}\" has unsupported type \"{1}\" Only strings ar allowed."
                                           , propinfo.Name, propinfo.PropertyType.FullName);
            }
        }

        public static List<XmlColumn> GetColumnsForType<T>() where T : new()
        {
            var bf = System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance;
            var item_type = typeof(T);
            var properties = item_type.GetProperties(bf);
            var target_props = properties.Where(p => p.CanRead).Where(p => p.PropertyType == typeof(string)).ToList();
            var cols = target_props.Select(p => new XmlColumn(p)).ToList();
            return cols;
        }
    }

}