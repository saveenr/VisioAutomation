using System.Collections.Generic;
using System.Linq;
using SXL = System.Xml.Linq;

namespace VisioAutomation.VDX.Internal.Extensions
{
    internal static class XMLExtensions
    {

        public static SXL.XElement ElementVisioSchema2003(this SXL.XElement el, string name)
        {
            string fullname = string.Format("{0}{1}", Constants.VisioXmlNamespace2003, name);
            var child_el = el.Element(fullname);
            return child_el;
        }

        public static IEnumerable<SXL.XElement> ElementsVisioSchema2003(this SXL.XElement el, string name)
        {
            string fullname = string.Format("{0}{1}", Constants.VisioXmlNamespace2003, name);
            var child_els = el.Elements(fullname);
            return child_els;
        }

        public static SXL.XElement RemoveElement(this SXL.XElement el, SXL.XName name)
        {
            var n = el.Element(name);
            if (n != null)
            {
                n.Remove();
                return n;
            }
            else
            {
                return null;
            }
        }

        public static SXL.XElement CleanElement(this SXL.XElement el, SXL.XName name)
        {
            var n = el.Element(name);
            if (n != null)
            {
                var nodes = n.Nodes().ToList();
                foreach (var c in nodes)
                {
                    c.Remove();
                }
                return n;
            }
            else
            {
                return null;
            }
        }

        public static void SetElementValueConditional(this SXL.XElement el, SXL.XName name, string value)
        {
            if (value != null)
            {
                el.SetElementValue(name, value);
            }
        }

        public static void SetElementValueConditional<T>(this SXL.XElement el, SXL.XName name, T? value) where T : struct
        {
            if (value != null)
            {
                el.SetElementValue(name, value.ToString());
            }
        }

        public static void SetElementValueConditional<T, TDest>(this SXL.XElement el, SXL.XName name, T? value, System.Func<T, TDest> xfrm) where T : struct
        {
            if (value != null)
            {
                var v = xfrm(value.Value);
                el.SetElementValue(name, v.ToString());
            }
        }

        public static void SetElementValueConditionalBool(this SXL.XElement el, SXL.XName name, bool? value)
        {
            if (value != null)
            {
                el.SetElementValue(name, value.Value ? "1" : "0");
            }
        }

        public static void SetElementValueConditionalDateTime(this SXL.XElement el, SXL.XName name, System.DateTimeOffset? date)
        {
            const string datefmt = "yyyy-MM-ddTHH:mm:ss";

            if (date!= null)
            {
                string datestr = date.Value.ToString(datefmt, System.Globalization.CultureInfo.InvariantCulture);
                el.SetElementValue( name, datestr);
            }
        }
    }
}