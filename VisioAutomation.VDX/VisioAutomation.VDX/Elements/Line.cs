using VisioAutomation.VDX.Internal;
using VA=VisioAutomation;
using SXL = System.Xml.Linq;

namespace VisioAutomation.VDX.Elements
{
    public class Line
    {
        public ShapeSheet.PointCell Weight = new ShapeSheet.PointCell();
        public ShapeSheet.ColorCell Color = new ShapeSheet.ColorCell();
        public ShapeSheet.IntCell Pattern = new ShapeSheet.IntCell();
        public ShapeSheet.DoubleCell Rounding = new ShapeSheet.DoubleCell();
        public ShapeSheet.IntCell EndArrowSize = new ShapeSheet.IntCell();
        public ShapeSheet.IntCell BeginArrowSize = new ShapeSheet.IntCell();
        public ShapeSheet.IntCell EndArrow = new ShapeSheet.IntCell();
        public ShapeSheet.IntCell BeginArrow = new ShapeSheet.IntCell();
        public ShapeSheet.IntCell Cap = new ShapeSheet.IntCell();
        public ShapeSheet.TransparencyCell Transparency = new ShapeSheet.TransparencyCell();

        public void AddToElement(SXL.XElement parent)
        {
            var line_el = XMLUtil.CreateVisioSchema2003Element("Line");
            line_el.Add(this.Weight.ToXml("LineWeight"));
            line_el.Add(this.Color.ToXml("LineColor"));
            line_el.Add(this.Pattern.ToXml("LinePattern"));
            line_el.Add(this.Rounding.ToXml("Rounding"));
            line_el.Add(this.EndArrowSize.ToXml("EndArrowSize"));
            line_el.Add(this.BeginArrowSize.ToXml("BeginArrowSize"));
            line_el.Add(this.EndArrow.ToXml("EndArrow"));
            line_el.Add(this.BeginArrow.ToXml("BeginArrow"));
            line_el.Add(this.Cap.ToXml("LineCap"));
            line_el.Add(this.Transparency.ToXml("LineColorTrans"));
            parent.Add(line_el);
        }
    }
}