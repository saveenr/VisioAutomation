using System.Collections.Generic;
using System.Linq;

namespace XmlPersist
{
    public class XmlTableWriter<T> where T : new()
    {
        public XmlTableWriter()
        {

        }

        public void SaveToFile(IEnumerable<T> items, string filename)
        {
            var xo = new System.Xml.XmlTextWriter(filename, System.Text.Encoding.UTF8);
            xo.Formatting = System.Xml.Formatting.Indented;
            xo.WriteStartDocument();

            var cols = XmlColumn.GetColumnsForType<T>();

            xo.WriteStartElement("table"); // <table>

            foreach (var item in items)
            {
                xo.WriteStartElement("row"); // <row>
                foreach (var col in cols)
                {
                    xo.WriteAttributeString(col.Name, col.GetStringValue(item));
                }
                xo.WriteEndElement(); // </row>

            }
            xo.WriteEndElement(); // </table>

            xo.WriteEndDocument();
            xo.Flush();
            xo.Close();
        }
    }

}