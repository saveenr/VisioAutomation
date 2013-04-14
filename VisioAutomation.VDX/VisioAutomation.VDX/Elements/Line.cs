using VisioAutomation.VDX.Internal;
using VA=VisioAutomation;
using SXL = System.Xml.Linq;

namespace VisioAutomation.VDX.Elements
{
    public class Line
    {
        public VA.VDX.ShapeSheet.PointCell Weight = new VA.VDX.ShapeSheet.PointCell();
        public VA.VDX.ShapeSheet.ColorCell Color = new VA.VDX.ShapeSheet.ColorCell();
        public VA.VDX.ShapeSheet.IntCell Pattern = new VA.VDX.ShapeSheet.IntCell();
        public VA.VDX.ShapeSheet.DoubleCell Rounding = new VA.VDX.ShapeSheet.DoubleCell();
        public VA.VDX.ShapeSheet.IntCell EndArrowSize = new VA.VDX.ShapeSheet.IntCell();
        public VA.VDX.ShapeSheet.IntCell BeginArrowSize = new VA.VDX.ShapeSheet.IntCell();
        public VA.VDX.ShapeSheet.IntCell EndArrow = new VA.VDX.ShapeSheet.IntCell();
        public VA.VDX.ShapeSheet.IntCell BeginArrow = new VA.VDX.ShapeSheet.IntCell();
        public VA.VDX.ShapeSheet.IntCell Cap = new VA.VDX.ShapeSheet.IntCell();
        public VA.VDX.ShapeSheet.TransparencyCell Transparency = new VA.VDX.ShapeSheet.TransparencyCell();

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