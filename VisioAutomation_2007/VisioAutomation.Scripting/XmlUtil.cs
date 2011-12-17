namespace VisioAutomation.Scripting
{
    static class XmlUtil
    {
        public static string GetAttributeValue(System.Xml.Linq.XElement el, System.Xml.Linq.XName name, string defval)
        {
            var attr = el.Attribute(name);
            if (attr == null)
            {
                return defval;
            }

            return attr.Value ?? defval;
        }

        public static T GetAttributeValue<T>(System.Xml.Linq.XElement el, System.Xml.Linq.XName name, System.Func<string, T> converter)
        {
            var a = el.Attribute(name);
            if (a == null)
            {
                string msg = string.Format(System.Globalization.CultureInfo.InvariantCulture, "Missing value for attribute \"{0}\"", name);
                throw new System.ArgumentException(msg);
            }
            string v = el.Attribute(name).Value;
            return converter(v);
        }

        public static T GetAttributeValue<T>(System.Xml.Linq.XElement el, System.Xml.Linq.XName name, T defval, System.Func<string, T> converter)
        {
            var a = el.Attribute(name);
            if (a == null)
            {
                return defval;
            }
            string v = el.Attribute(name).Value;
            return converter(v);
        }
    }
}