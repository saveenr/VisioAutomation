using VA = VisioAutomation;

namespace VisioAutomation.Scripting.DirectedGraph
{
    public static class XmlExtensions
    {
        public static VA.Drawing.ColorRGB AttributeAsColor(this System.Xml.Linq.XElement el, string name,
                                                     VA.Drawing.ColorRGB def)
        {
            return VA.Scripting.XmlUtil.GetAttributeValue(el, name, def, s => VA.Drawing.ColorRGB.ParseWebColor(s));
        }

        public static double AttributeAsInches(this System.Xml.Linq.XElement el, string name, double def)
        {
            return VA.Scripting.XmlUtil.GetAttributeValue(el, name, def, s => PointsToInches(double.Parse(s)));
        }

        private static double PointsToInches(double points)
        {
            return points / 72.0;
        }
    }
}