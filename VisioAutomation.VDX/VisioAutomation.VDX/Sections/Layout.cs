using VisioAutomation.VDX.Enums;
using VisioAutomation.VDX.Internal;
using VisioAutomation.VDX.ShapeSheet;
using SXL = System.Xml.Linq;

namespace VisioAutomation.VDX.Sections
{
    public class Layout
    {
        public BoolCell ShapePermeableX = new BoolCell();
        public BoolCell ShapePermeableY = new BoolCell();
        public BoolCell ShapePermeablePlace = new BoolCell();
        public EnumCell<ShapeFixedCodeType> ShapeFixedCode = new EnumCell<ShapeFixedCodeType>(v => (int) v);
        public EnumCell<ShapePlowCodeType> ShapePlowCode = new EnumCell<ShapePlowCodeType>(v => (int) v);
        public EnumCell<RouteStyle> ShapeRouteStyle = new EnumCell<RouteStyle>(v => (int) v);
        public EnumCell<ConFixedCode> ConFixedCode = new EnumCell<ConFixedCode>(v => (int) v);
        public EnumCell<ConLineJumpCode> ConLineJumpCode = new EnumCell<ConLineJumpCode>(v => (int) v);
        public EnumCell<ConLineJumpStyle> ConLineJumpStyle = new EnumCell<ConLineJumpStyle>(v => (int) v);
        public EnumCell<ConLineJumpDirX> ConLineJumpDirX = new EnumCell<ConLineJumpDirX>(v => (int) v);
        public EnumCell<ConLineJumpDirY> ConLineJumpDirY = new EnumCell<ConLineJumpDirY>(v => (int) v);
        public EnumCell<PlaceFlip> ShapePlaceFlip = new EnumCell<PlaceFlip>(v => (int) v);
        public EnumCell<ConLineRouteExt> ConLineRouteExt = new EnumCell<ConLineRouteExt>(v => (int) v);
        public EnumCell<PageShapeSplit> ShapeSplit = new EnumCell<PageShapeSplit>(v => (int) v);

        public EnumCell<ShapeSplittable> ShapeSplittable = new EnumCell<ShapeSplittable>(v => (int) v);

        public void AddToElement(SXL.XElement parent)
        {
            var el = XMLUtil.CreateVisioSchema2003Element("Layout");

            el.Add(this.ShapePermeableX.ToXml("ShapePermeableX"));
            el.Add(this.ShapePermeableY.ToXml("ShapePermeableY"));
            el.Add(this.ShapePermeablePlace.ToXml("ShapePermeablePlace"));
            el.Add(this.ShapeFixedCode.ToXml("ShapeFixedCode"));
            el.Add(this.ShapePlowCode.ToXml("ShapePlowCode"));
            el.Add(this.ShapeRouteStyle.ToXml("ShapeRouteStyle"));
            el.Add(this.ConFixedCode.ToXml("ConFixedCode"));
            el.Add(this.ConLineJumpCode.ToXml("ConLineJumpCode"));
            el.Add(this.ConLineJumpStyle.ToXml("ConLineJumpStyle"));
            el.Add(this.ConLineJumpDirX.ToXml("ConLineJumpDirX"));
            el.Add(this.ConLineJumpDirY.ToXml("ConLineJumpDirY"));
            el.Add(this.ShapePlaceFlip.ToXml("ShapePlaceFlip"));
            el.Add(this.ConLineRouteExt.ToXml("ConLineRouteExt"));

            el.Add(this.ShapeSplit.ToXml("ShapeSplit"));
            el.Add(this.ShapeSplittable.ToXml("ShapeSplittable"));

            parent.Add(el);
        }
    }
}