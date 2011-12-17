using VA = VisioAutomation;

namespace VisioAutomation.Scripting.FlowChart
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
            return VA.Scripting.XmlUtil.GetAttributeValue(el, name, def, s => Convert.PointsToInches(double.Parse(s)));
        }
    }
}