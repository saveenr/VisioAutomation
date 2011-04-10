using System.Linq;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using System.Collections.Generic;


namespace VisioAutomation.Layout
{

    public class ShapeLayoutQuery : VA.ShapeSheet.Query.CellQuery
    {
        public VA.ShapeSheet.Query.CellQueryColumn ConFixedCode { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn ConLineJumpCode { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn ConLineJumpDirX { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn ConLineJumpDirY { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn ConLineJumpStyle { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn ConLineRouteExt { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn ShapeFixedCode { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn ShapePermeablePlace { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn ShapePermeableX { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn ShapePermeableY { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn ShapePlaceFlip { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn ShapePlaceStyle { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn ShapePlowCode { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn ShapeRouteStyle { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn ShapeSplit { get; set; }
        public VA.ShapeSheet.Query.CellQueryColumn ShapeSplittable { get; set; }

        public ShapeLayoutQuery() :
            base()
        {
            this.ConFixedCode = this.AddColumn(VA.ShapeSheet.SRCConstants.ConFixedCode, "ConFixedCode");
            this.ConLineJumpCode = this.AddColumn(VA.ShapeSheet.SRCConstants.ConLineJumpCode, "ConLineJumpCode");
            this.ConLineJumpDirX = this.AddColumn(VA.ShapeSheet.SRCConstants.ConLineJumpDirX, "ConLineJumpDirX");
            this.ConLineJumpDirY = this.AddColumn(VA.ShapeSheet.SRCConstants.ConLineJumpDirY, "ConLineJumpDirY");
            this.ConLineJumpStyle = this.AddColumn(VA.ShapeSheet.SRCConstants.ConLineJumpStyle, "ConLineJumpStyle");
            this.ConLineRouteExt = this.AddColumn(VA.ShapeSheet.SRCConstants.ConLineRouteExt, "ConLineRouteExt");
            this.ShapeFixedCode = this.AddColumn(VA.ShapeSheet.SRCConstants.ShapeFixedCode, "ShapeFixedCode");
            this.ShapePermeablePlace = this.AddColumn(VA.ShapeSheet.SRCConstants.ShapePermeablePlace, "ShapePermeablePlace");
            this.ShapePermeableX = this.AddColumn(VA.ShapeSheet.SRCConstants.ShapePermeableX, "ShapePermeableX");
            this.ShapePermeableY = this.AddColumn(VA.ShapeSheet.SRCConstants.ShapePermeableY, "ShapePermeableY");
            this.ShapePlaceFlip = this.AddColumn(VA.ShapeSheet.SRCConstants.ShapePlaceFlip, "ShapePlaceFlip");
            this.ShapePlaceStyle = this.AddColumn(VA.ShapeSheet.SRCConstants.ShapePlaceStyle, "ShapePlaceStyle");
            this.ShapePlowCode = this.AddColumn(VA.ShapeSheet.SRCConstants.ShapePlowCode, "ShapePlowCode");
            this.ShapeRouteStyle = this.AddColumn(VA.ShapeSheet.SRCConstants.ShapeRouteStyle, "ShapeRouteStyle");
            this.ShapeSplit = this.AddColumn(VA.ShapeSheet.SRCConstants.ShapeSplit, "ShapeSplit");
            this.ShapeSplittable = this.AddColumn(VA.ShapeSheet.SRCConstants.ShapeSplittable, "ShapeSplittable");
        }

    }
}