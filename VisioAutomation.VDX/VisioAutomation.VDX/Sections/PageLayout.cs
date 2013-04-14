using VisioAutomation.VDX.Internal;
using VisioAutomation.VDX.ShapeSheet;
using SXL = System.Xml.Linq;

namespace VisioAutomation.VDX.Sections
{
    public class PageLayout
    {
        public BoolCell ResizePage = new BoolCell();
        public BoolCell EnableGrid = new BoolCell();
        public BoolCell DynamicsOff = new BoolCell();
        public IntCell PlaceStyle = new IntCell();

        public IntCell RouteStyle = new IntCell();
        public DoubleCell PlaceDepth = new DoubleCell();
        public IntCell PlowCode = new IntCell();

        public IntCell LineJumpCode = new IntCell();
        public IntCell LineJumpStyle = new IntCell();

        public IntCell PageLineJumpDirX = new IntCell();
        public IntCell PageLineJumpDirY = new IntCell();

        public IntCell LineToNodeX = new IntCell();
        public IntCell LineToNodeY = new IntCell();

        public IntCell BlockSizeX = new IntCell();
        public IntCell BlockSizeY = new IntCell();

        public IntCell AvenueSizeX = new IntCell();
        public IntCell AvenueSizeY = new IntCell();

        public IntCell LineToLineX = new IntCell();
        public IntCell LineToLineY = new IntCell();

        public IntCell LineJumpFactorX = new IntCell();
        public IntCell LineJumpFactorY = new IntCell();

        public IntCell LineAdjustFrom = new IntCell();
        public IntCell LineAdjustTo = new IntCell();

        public IntCell PlaceFlip = new IntCell();
        public IntCell LineRouteExt = new IntCell();
        public IntCell PageShapeSplit = new IntCell();

        public void AddToElement(SXL.XElement parent)
        {
            var el = XMLUtil.CreateVisioSchema2003Element("PageLayout");
            el.Add(this.ResizePage.ToXml("ResizePage"));
            el.Add(this.EnableGrid.ToXml("EnableGrid"));
            el.Add(this.DynamicsOff.ToXml("DynamicsOff"));
            el.Add(this.PlaceStyle.ToXml("PlaceStyle"));
            el.Add(this.RouteStyle.ToXml("RouteStyle"));
            el.Add(this.PlaceDepth.ToXml("PlaceDepth"));
            el.Add(this.PlowCode.ToXml("PlowCode"));
            el.Add(this.LineJumpCode.ToXml("LineJumpCode"));
            el.Add(this.LineJumpStyle.ToXml("LineJumpStyle"));
            el.Add(this.PageLineJumpDirX.ToXml("PageLineJumpDirX"));
            el.Add(this.PageLineJumpDirY.ToXml("PageLineJumpDirY"));
            el.Add(this.LineToNodeX.ToXml("LineToNodeX"));
            el.Add(this.LineToNodeY.ToXml("LineToNodeY"));
            el.Add(this.BlockSizeX.ToXml("BlockSizeX"));
            el.Add(this.BlockSizeY.ToXml("BlockSizeY"));
            el.Add(this.AvenueSizeX.ToXml("AvenueSizeX"));
            el.Add(this.AvenueSizeY.ToXml("AvenueSizeY"));
            el.Add(this.LineToLineX.ToXml("LineToLineX"));
            el.Add(this.LineToLineY.ToXml("LineToLineY"));
            el.Add(this.LineJumpFactorX.ToXml("LineJumpFactorX"));
            el.Add(this.LineJumpFactorY.ToXml("LineJumpFactorY"));
            el.Add(this.LineAdjustFrom.ToXml("LineAdjustFrom"));
            el.Add(this.LineAdjustTo.ToXml("LineAdjustTo"));

            el.Add(this.PlaceFlip.ToXml("PlaceFlip"));
            el.Add(this.LineRouteExt.ToXml("LineRouteExt"));
            el.Add(this.PageShapeSplit.ToXml("PageShapeSplit"));

            parent.Add(el);
        }
    }
}