using SXL = System.Xml.Linq;

namespace VisioAutomation.Scripting
{
    internal static class XmlUtil
    {
        public static string GetAttributeValue(SXL.XElement el, SXL.XName name, string defval)
        {
            var attr = el.Attribute(name);
            if (attr == null)
            {
                return defval;
            }

            return attr.Value;
        }

        public static T GetAttributeValue<T>(SXL.XElement el, SXL.XName name, System.Func<string, T> converter)
        {
            var a = el.Attribute(name);
            if (a == null)
            {
                var culture = System.Globalization.CultureInfo.InvariantCulture;
                string msg = string.Format(culture, "Missing value for attribute \"{0}\"", name);
                throw new System.ArgumentException(msg);
            }

            string v = a.Value;
            return converter(v);                
        }

        public static T GetAttributeValue<T>(SXL.XElement el, SXL.XName name, T defval, System.Func<string, T> converter)
        {
            var a = el.Attribute(name);
            if (a == null)
            {
                return defval;
            }

            string v = a.Value;
            return converter(v);
        }
    }
}