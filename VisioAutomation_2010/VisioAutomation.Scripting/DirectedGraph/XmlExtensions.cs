using VisioAutomation.Colors;
using VisioAutomation.Scripting.Utilities;
using SXL = System.Xml.Linq;

namespace VisioAutomation.Scripting.DirectedGraph
{
    static class XmlExtensions
    {
        public static ColorRGB AttributeAsColor(this SXL.XElement el, string name,
                                                     ColorRGB def)
        {
            return XmlUtil.GetAttributeValue(el, name, def, ColorRGB.ParseWebColor);
        }

        public static double AttributeAsInches(this SXL.XElement el, string name, double def)
        {
            return XmlUtil.GetAttributeValue(el, name, def, s => XmlExtensions.PointsToInches(double.Parse(s)));
        }

        private static double PointsToInches(double points)
        {
            return points / 72.0;
        }
    }
}