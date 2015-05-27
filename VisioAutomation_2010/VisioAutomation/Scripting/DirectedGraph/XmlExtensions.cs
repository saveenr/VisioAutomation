using SXL = System.Xml.Linq;

namespace VisioAutomation.Scripting.DirectedGraph
{
    public static class XmlExtensions
    {
        public static Drawing.ColorRGB AttributeAsColor(this SXL.XElement el, string name,
                                                     Drawing.ColorRGB def)
        {
            return XmlUtil.GetAttributeValue(el, name, def, Drawing.ColorRGB.ParseWebColor);
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