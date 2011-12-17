using VisioAutomation.VDX.Internal;
using VisioAutomation.VDX.ShapeSheet;

namespace VisioAutomation.VDX.Sections
{
    public class GeomArcTo : GeomRow
    {
        public DistanceCell X = new DistanceCell();
        public DistanceCell Y = new DistanceCell();
        public StringCell A = new StringCell();

        public override void AddToElement(System.Xml.Linq.XElement parent, int index)
        {
            var el = XMLUtil.CreateVisioSchema2003Element("ArcTo");
            el.SetAttributeValue("IX", index.ToString(System.Globalization.CultureInfo.InvariantCulture));
            el.Add(this.X.ToXml("X"));
            el.Add(this.Y.ToXml("Y"));
            el.Add(this.A.ToXml("A"));
            parent.Add(el);
        }
    }
}