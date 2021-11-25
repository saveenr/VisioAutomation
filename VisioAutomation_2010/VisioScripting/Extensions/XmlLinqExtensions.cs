

namespace VisioScripting.Extensions
{
    static class XmlLinqExtensions
    {
        public static VisioAutomation.Models.Color.ColorRgb AttributeAsColor(this SXL.XElement el, string name,
            VisioAutomation.Models.Color.ColorRgb def)
        {
            return el.GetAttributeValue(name, def, VisioAutomation.Models.Color.ColorRgb.ParseWebColor);
        }

        public static double AttributeAsInches(this SXL.XElement el, string name, double def)
        {
            var culture = System.Globalization.CultureInfo.InvariantCulture;
            double DoubleParse(string str) => double.Parse(str, culture);
            return el.GetAttributeValue(name, def, s => XmlLinqExtensions._points_to_inches(DoubleParse(s)));
        }

        private static double _points_to_inches(double points)
        {
            return points/72.0;
        }

        public static string GetAttributeValue(this SXL.XElement el, SXL.XName name, string defval)
        {
            var attr = el.Attribute(name);
            if (attr == null)
            {
                return defval;
            }

            return attr.Value;
        }

        public static T GetAttributeValue<T>(this SXL.XElement el, SXL.XName name, System.Func<string, T> converter)
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

        public static T GetAttributeValue<T>(this SXL.XElement el, SXL.XName name, T defval, System.Func<string, T> converter)
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