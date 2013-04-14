using VisioAutomation.VDX.Internal;
using VisioAutomation.VDX.ShapeSheet;
using SXL = System.Xml.Linq;

namespace VisioAutomation.VDX.Sections
{
    public class Fill
    {
        public ColorCell ForegroundColor = new ColorCell();
        public ColorCell BackgroundColor = new ColorCell();
        public IntCell Pattern = new IntCell();
        public ColorCell ShadowForegroundColor = new ColorCell();
        public ColorCell ShadowBackgroundColor = new ColorCell();
        public IntCell ShadowPattern = new IntCell();
        public TransparencyCell ForegroundTransparency = new TransparencyCell();
        public TransparencyCell BackgroundTransparency = new TransparencyCell();
        public TransparencyCell ShadowForegroundTransparency = new TransparencyCell();
        public TransparencyCell ShadowBackgroundTransparency = new TransparencyCell();

        public IntCell ShadowType = new IntCell();
        public DistanceCell ShadowOffsetX = new DistanceCell();
        public DistanceCell ShadowOffsetY = new DistanceCell();
        public AngleCell ShadowObliqueAngle = new AngleCell();
        public DoubleCell ShadowScale = new DoubleCell();

        public void AddToElement(SXL.XElement parent)
        {
            var el = XMLUtil.CreateVisioSchema2003Element("Fill");
            el.Add(this.ForegroundColor.ToXml("FillForegnd"));
            el.Add(this.BackgroundColor.ToXml("FillBkgnd"));
            el.Add(this.Pattern.ToXml("FillPattern"));
            el.Add(this.ShadowForegroundColor.ToXml("ShdwForegnd"));
            el.Add(this.ShadowBackgroundColor.ToXml("ShdwBkgnd"));
            el.Add(this.ShadowPattern.ToXml("ShdwPattern"));

            el.Add(this.ForegroundTransparency.ToXml("FillForegndTrans"));
            el.Add(this.BackgroundTransparency.ToXml("FillBkgndTrans"));

            el.Add(
                this.ShadowForegroundTransparency.ToXml("ShdwForegndTrans"));
            el.Add(
                this.ShadowBackgroundTransparency.ToXml("ShdwBkgndTrans"));

            el.Add(this.ShadowType.ToXml("ShapeShdwType"));
            el.Add(this.ShadowOffsetX.ToXml("ShapeShdwOffsetX"));
            el.Add(this.ShadowOffsetY.ToXml("ShapeShdwOffsetY"));
            el.Add(this.ShadowObliqueAngle.ToXml("ShapeShdwObliqueAngle"));
            el.Add(this.ShadowScale.ToXml("ShapeShdwScaleFactor"));

            parent.Add(el);
        }
    }
}