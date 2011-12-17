using System.Collections.Generic;
using System.Linq;

namespace XmlPersist
{
    public class XmlTableReader<T> where T : new()
    {
        public XmlTableReader()
        {

        }

        public IEnumerable<T> LoadFromFile(string filename)
        {
            var doc = System.Xml.Linq.XDocument.Load(filename);
            return Load(doc);
        }

        public IEnumerable<T> LoadFromString(string text)
        {
            var doc = System.Xml.Linq.XDocument.Parse(text);
            return Load(doc);
        }

        public IEnumerable<T> Load(System.Xml.Linq.XDocument doc)
        {
            var cols = XmlColumn.GetColumnsForType<T>();

            var root_el = doc.Root;
            foreach (var row_el in root_el.Elements("row"))
            {
                var new_item = new T();
                foreach (var col in cols)
                {
                    var attr = row_el.Attribute(col.Name);
                    if (attr == null)
                    {
                        // do nothing
                        continue;
                    }

                    if (col.PropertyInfo.PropertyType != typeof (string))
                    {
                        string msg =
                            string.Format(
                                "class {0} has unsupported property datatype. Property {1} has datatype {2}.",
                                typeof (T).FullName, col.PropertyInfo.Name, col.PropertyInfo.PropertyType.FullName);
                        throw new System.Exception(msg);
                    }

                    string prop_string_val = attr.Value;
                    col.SetStringValue(new_item, prop_string_val);
                }
                yield return new_item;
            }
        }
    }
}
