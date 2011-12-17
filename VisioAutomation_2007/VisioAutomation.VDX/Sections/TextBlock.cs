using VisioAutomation.VDX.Internal;
using VisioAutomation.VDX.ShapeSheet;

namespace VisioAutomation.VDX.Sections
{
    public class TextBlock
    {
        public DistanceCell LeftMargin = new DistanceCell();
        public DistanceCell RightMargin = new DistanceCell();
        public DistanceCell TopMargin = new DistanceCell();
        public DistanceCell BottomMargin = new DistanceCell();

        public IntCell VerticalAlign = new IntCell();
        public ColorCell TextBkgnd = new ColorCell();

        public DistanceCell DefaultTabStop = new DistanceCell();
        public IntCell TextDirection = new IntCell();
        public TransparencyCell TextBkgndTrans = new TransparencyCell();

        public void AddToElement(System.Xml.Linq.XElement parent)
        {
            var el1 = XMLUtil.CreateVisioSchema2003Element("TextBlock");
            el1.Add(this.LeftMargin.ToXml("LeftMargin"));
            el1.Add(this.RightMargin.ToXml("RightMargin"));
            el1.Add(this.TopMargin.ToXml("TopMargin"));
            el1.Add(this.BottomMargin.ToXml("BottomMargin"));
            el1.Add(this.VerticalAlign.ToXml("VerticalAlign"));
            el1.Add(this.TextBkgnd.ToXml("TextBkgnd"));

            el1.Add(this.DefaultTabStop.ToXml("DefaultTabStop"));
            el1.Add(this.TextDirection.ToXml("TextDirection"));
            el1.Add(this.TextBkgndTrans.ToXml("TextBkgndTrans"));

            parent.Add(el1);
        }
    }
}