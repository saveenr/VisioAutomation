using VA = VisioAutomation;
using SXL = System.Xml.Linq;

namespace VisioAutomation.Scripting.DirectedGraph
{
    public static class XmlExtensions
    {
        public static VA.Drawing.ColorRGB AttributeAsColor(this SXL.XElement el, string name,
                                                     VA.Drawing.ColorRGB def)
        {
            return VA.Scripting.XmlUtil.GetAttributeValue(el, name, def, VA.Drawing.ColorRGB.ParseWebColor);
        }

        public static double AttributeAsInches(this SXL.XElement el, string name, double def)
        {
            return VA.Scripting.XmlUtil.GetAttributeValue(el, name, def, s => PointsToInches(double.Parse(s)));
        }

        private static double PointsToInches(double points)
        {
            return points / 72.0;
        }
    }
}