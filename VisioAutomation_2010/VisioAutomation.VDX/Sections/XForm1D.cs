using VisioAutomation.VDX.Internal;
using VisioAutomation.VDX.ShapeSheet;

namespace VisioAutomation.VDX.Sections
{
    public class XForm1D
    {
        public DistanceCell BeginX = new DistanceCell();
        public DistanceCell BeginY = new DistanceCell();
        public DistanceCell EndX = new DistanceCell();
        public DistanceCell EndY = new DistanceCell();

        public void AddToElement(System.Xml.Linq.XElement parent)
        {
            var el = XMLUtil.CreateVisioSchema2003Element("XForm1D");
            el.Add(this.BeginX.ToXml("BeginX"));
            el.Add(this.BeginY.ToXml("BeginY"));
            el.Add(this.EndX.ToXml("EndX"));
            el.Add(this.EndY.ToXml("EndY"));

            parent.Add(el);
        }
    }
}