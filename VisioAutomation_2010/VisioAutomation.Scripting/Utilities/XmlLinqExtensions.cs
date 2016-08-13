using VisioAutomation.Colors;
using SXL = System.Xml.Linq;

namespace VisioAutomation.Scripting.Utilities
{
    static class XmlLinqExtensions
    {
        public static ColorRGB AttributeAsColor(this SXL.XElement el, string name,
            ColorRGB def)
        {
            return el.GetAttributeValue(name, def, ColorRGB.ParseWebColor);
        }

        public static double AttributeAsInches(this SXL.XElement el, string name, double def)
        {
            return el.GetAttributeValue(name, def, s => XmlLinqExtensions.PointsToInches(double.Parse(s)));
        }

        private static double PointsToInches(double points)
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