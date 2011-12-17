using VisioAutomation.VDX.Internal;
using VisioAutomation.VDX.ShapeSheet;

namespace VisioAutomation.VDX.Sections
{
    public class LineTo : GeomRow
    {
        public DistanceCell X = new DistanceCell();
        public DistanceCell Y = new DistanceCell();

        public override void AddToElement(System.Xml.Linq.XElement parent, int index)
        {
            var el = XMLUtil.CreateVisioSchema2003Element("LineTo");
            el.SetAttributeValue("IX", index.ToString(System.Globalization.CultureInfo.InvariantCulture));
            el.Add(this.X.ToXml("X"));
            el.Add(this.Y.ToXml("Y"));
            parent.Add(el);
        }

        public LineTo(double x, double y)
        {
            this.X.Result = x;
            this.Y.Result = y;
        }
    }
}