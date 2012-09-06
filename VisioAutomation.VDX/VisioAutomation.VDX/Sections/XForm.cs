using VisioAutomation.VDX.Internal;
using VisioAutomation.VDX.ShapeSheet;

namespace VisioAutomation.VDX.Sections
{
    public class XForm
    {
        public DistanceCell PinX = new DistanceCell();
        public DistanceCell PinY = new DistanceCell();
        public DistanceCell Width = new DistanceCell();
        public DistanceCell Height = new DistanceCell();
        public DistanceCell LocPinX = new DistanceCell();
        public DistanceCell LocPinY = new DistanceCell();
        public AngleCell Angle = new AngleCell();
        public IntCell FlipX = new IntCell();
        public IntCell FlipY = new IntCell();
        public IntCell FlipMode = new IntCell();

        public void AddToElement(System.Xml.Linq.XElement parent)
        {
            var el = XMLUtil.CreateVisioSchema2003Element("XForm");
            el.Add(this.PinX.ToXml("PinX"));
            el.Add(this.PinY.ToXml("PinY"));
            el.Add(this.Width.ToXml("Width"));
            el.Add(this.Height.ToXml("Height"));
            el.Add(this.LocPinX.ToXml("LocPinX"));
            el.Add(this.LocPinY.ToXml("LocPinY"));
            el.Add(this.Angle.ToXml("Angle"));
            el.Add(this.FlipX.ToXml("FlipX"));
            el.Add(this.FlipY.ToXml("FlipY"));
            el.Add(this.FlipMode.ToXml("FlipMode"));

            parent.Add(el);
        }
    }
}