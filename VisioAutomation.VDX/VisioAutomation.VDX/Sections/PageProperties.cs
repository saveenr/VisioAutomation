using VisioAutomation.VDX.Internal;
using VisioAutomation.VDX.ShapeSheet;
using SXL = System.Xml.Linq;

namespace VisioAutomation.VDX.Sections
{
    public class PageProperties
    {
        public DistanceCell PageWidth = new DistanceCell();
        public DistanceCell PageHeight = new DistanceCell();
        public DistanceCell ShdwOffsetX = new DistanceCell();
        public DistanceCell ShdwOffsetY = new DistanceCell();
        public DistanceCell PageScale = new DistanceCell();
        public IntCell DrawingSizeType = new IntCell();
        public IntCell DrawingScaleType = new IntCell();
        public IntCell InhibitSnap = new IntCell();
        public IntCell UIVisibility = new IntCell();
        public IntCell ShdwType = new IntCell();

        public IntCell ShdwObliqueAngle = new IntCell();
        public IntCell ShdwScaleFactor = new IntCell();

        public void AddToElement(SXL.XElement parent)
        {
            var el = XMLUtil.CreateVisioSchema2003Element("PageProps");
            el.Add(this.PageWidth.ToXml("PageWidth"));
            el.Add(this.PageHeight.ToXml("PageHeight"));
            el.Add(this.ShdwOffsetX.ToXml("ShdwOffsetX"));
            el.Add(this.ShdwOffsetY.ToXml("ShdwOffsetY"));
            el.Add(this.PageScale.ToXml("PageScale"));
            el.Add(this.DrawingSizeType.ToXml("DrawingSizeType"));
            el.Add(this.DrawingScaleType.ToXml("DrawingScaleType"));
            el.Add(this.InhibitSnap.ToXml("InhibitSnap"));
            el.Add(this.UIVisibility.ToXml("UIVisibility"));
            el.Add(this.ShdwType.ToXml("ShdwType"));
            el.Add(this.ShdwObliqueAngle.ToXml("ShdwObliqueAngle"));
            el.Add(this.ShdwScaleFactor.ToXml("ShdwScaleFactor"));
            parent.Add(el);
        }
    }
}