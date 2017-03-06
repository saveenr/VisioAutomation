using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Shapes.Layout
{
    class ShapeLayoutCellsReader : SingleRowReader<Shapes.Layout.ShapeLayoutCells>
    {
        public CellColumn ConFixedCode { get; set; }
        public CellColumn ConLineJumpCode { get; set; }
        public CellColumn ConLineJumpDirX { get; set; }
        public CellColumn ConLineJumpDirY { get; set; }
        public CellColumn ConLineJumpStyle { get; set; }
        public CellColumn ConLineRouteExt { get; set; }
        public CellColumn ShapeFixedCode { get; set; }
        public CellColumn ShapePermeablePlace { get; set; }
        public CellColumn ShapePermeableX { get; set; }
        public CellColumn ShapePermeableY { get; set; }
        public CellColumn ShapePlaceFlip { get; set; }
        public CellColumn ShapePlaceStyle { get; set; }
        public CellColumn ShapePlowCode { get; set; }
        public CellColumn ShapeRouteStyle { get; set; }
        public CellColumn ShapeSplit { get; set; }
        public CellColumn ShapeSplittable { get; set; }
        public CellColumn DisplayLevel { get; set; }
        public CellColumn Relationships { get; set; }

        public ShapeLayoutCellsReader() 
        {
            this.ConFixedCode = this.query.AddCell(SrcConstants.ShapeLayoutConFixedCode, nameof(SrcConstants.ShapeLayoutConFixedCode));
            this.ConLineJumpCode = this.query.AddCell(SrcConstants.ShapeLayoutConLineJumpCode, nameof(SrcConstants.ShapeLayoutConLineJumpCode));
            this.ConLineJumpDirX = this.query.AddCell(SrcConstants.ShapeLayoutConLineJumpDirX, nameof(SrcConstants.ShapeLayoutConLineJumpDirX));
            this.ConLineJumpDirY = this.query.AddCell(SrcConstants.ShapeLayoutConLineJumpDirY, nameof(SrcConstants.ShapeLayoutConLineJumpDirY));
            this.ConLineJumpStyle = this.query.AddCell(SrcConstants.ShapeLayoutConLineJumpStyle, nameof(SrcConstants.ShapeLayoutConLineJumpStyle));
            this.ConLineRouteExt = this.query.AddCell(SrcConstants.ShapeLayoutConLineRouteExt, nameof(SrcConstants.ShapeLayoutConLineRouteExt));
            this.ShapeFixedCode = this.query.AddCell(SrcConstants.ShapeLayoutFixedCode, nameof(SrcConstants.ShapeLayoutFixedCode));
            this.ShapePermeablePlace = this.query.AddCell(SrcConstants.ShapeLayoutPermeablePlace, nameof(SrcConstants.ShapeLayoutPermeablePlace));
            this.ShapePermeableX = this.query.AddCell(SrcConstants.ShapeLayoutPermeableX, nameof(SrcConstants.ShapeLayoutPermeableX));
            this.ShapePermeableY = this.query.AddCell(SrcConstants.ShapeLayoutPermeableY, nameof(SrcConstants.ShapeLayoutPermeableY));
            this.ShapePlaceFlip = this.query.AddCell(SrcConstants.ShapeLayoutPlaceFlip, nameof(SrcConstants.ShapeLayoutPlaceFlip));
            this.ShapePlaceStyle = this.query.AddCell(SrcConstants.ShapeLayoutPlaceStyle, nameof(SrcConstants.ShapeLayoutPlaceStyle));
            this.ShapePlowCode = this.query.AddCell(SrcConstants.ShapeLayoutPlowCode, nameof(SrcConstants.ShapeLayoutPlowCode));
            this.ShapeRouteStyle = this.query.AddCell(SrcConstants.ShapeLayoutRouteStyle, nameof(SrcConstants.ShapeLayoutRouteStyle));
            this.ShapeSplit = this.query.AddCell(SrcConstants.ShapeLayoutSplit, nameof(SrcConstants.ShapeLayoutSplit));
            this.ShapeSplittable = this.query.AddCell(SrcConstants.ShapeLayoutSplittable, nameof(SrcConstants.ShapeLayoutSplittable));
            this.DisplayLevel = this.query.AddCell(SrcConstants.ShapeLayoutDisplayLevel, nameof(SrcConstants.ShapeLayoutDisplayLevel));
            this.Relationships = this.query.AddCell(SrcConstants.ShapeLayoutRelationships, nameof(SrcConstants.ShapeLayoutRelationships));


        }

        public override Shapes.Layout.ShapeLayoutCells CellDataToCellGroup(VisioAutomation.Utilities.ArraySegment<ShapeSheet.CellData> row)
        {
            var cells = new Shapes.Layout.ShapeLayoutCells();
            cells.ConFixedCode = row[this.ConFixedCode];
            cells.ConLineJumpCode = row[this.ConLineJumpCode];
            cells.ConLineJumpDirX = row[this.ConLineJumpDirX];
            cells.ConLineJumpDirY = row[this.ConLineJumpDirY];
            cells.ConLineJumpStyle = row[this.ConLineJumpStyle];
            cells.ConLineRouteExt = row[this.ConLineRouteExt];
            cells.FixedCode = row[this.ShapeFixedCode];
            cells.PermeablePlace = row[this.ShapePermeablePlace];
            cells.PermeableX = row[this.ShapePermeableX];
            cells.PermeableY = row[this.ShapePermeableY];
            cells.PlaceFlip = row[this.ShapePlaceFlip];
            cells.PlaceStyle = row[this.ShapePlaceStyle];
            cells.PlowCode = row[this.ShapePlowCode];
            cells.RouteStyle = row[this.ShapeRouteStyle];
            cells.Split = row[this.ShapeSplit];
            cells.Splittable = row[this.ShapeSplittable];
            cells.DisplayLevel = row[this.DisplayLevel];
            cells.Relationships = row[this.Relationships];
            return cells;
        }
    }
}