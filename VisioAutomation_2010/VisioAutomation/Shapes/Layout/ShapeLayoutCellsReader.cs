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
            this.ConFixedCode = this.query.AddCell(SrcConstants.ConFixedCode, nameof(SrcConstants.ConFixedCode));
            this.ConLineJumpCode = this.query.AddCell(SrcConstants.ConLineJumpCode, nameof(SrcConstants.ConLineJumpCode));
            this.ConLineJumpDirX = this.query.AddCell(SrcConstants.ConLineJumpDirX, nameof(SrcConstants.ConLineJumpDirX));
            this.ConLineJumpDirY = this.query.AddCell(SrcConstants.ConLineJumpDirY, nameof(SrcConstants.ConLineJumpDirY));
            this.ConLineJumpStyle = this.query.AddCell(SrcConstants.ConLineJumpStyle, nameof(SrcConstants.ConLineJumpStyle));
            this.ConLineRouteExt = this.query.AddCell(SrcConstants.ConLineRouteExt, nameof(SrcConstants.ConLineRouteExt));
            this.ShapeFixedCode = this.query.AddCell(SrcConstants.ShapeFixedCode, nameof(SrcConstants.ShapeFixedCode));
            this.ShapePermeablePlace = this.query.AddCell(SrcConstants.ShapePermeablePlace, nameof(SrcConstants.ShapePermeablePlace));
            this.ShapePermeableX = this.query.AddCell(SrcConstants.ShapePermeableX, nameof(SrcConstants.ShapePermeableX));
            this.ShapePermeableY = this.query.AddCell(SrcConstants.ShapePermeableY, nameof(SrcConstants.ShapePermeableY));
            this.ShapePlaceFlip = this.query.AddCell(SrcConstants.ShapePlaceFlip, nameof(SrcConstants.ShapePlaceFlip));
            this.ShapePlaceStyle = this.query.AddCell(SrcConstants.ShapePlaceStyle, nameof(SrcConstants.ShapePlaceStyle));
            this.ShapePlowCode = this.query.AddCell(SrcConstants.ShapePlowCode, nameof(SrcConstants.ShapePlowCode));
            this.ShapeRouteStyle = this.query.AddCell(SrcConstants.ShapeRouteStyle, nameof(SrcConstants.ShapeRouteStyle));
            this.ShapeSplit = this.query.AddCell(SrcConstants.ShapeSplit, nameof(SrcConstants.ShapeSplit));
            this.ShapeSplittable = this.query.AddCell(SrcConstants.ShapeSplittable, nameof(SrcConstants.ShapeSplittable));
            this.DisplayLevel = this.query.AddCell(SrcConstants.DisplayLevel, nameof(SrcConstants.DisplayLevel));
            this.Relationships = this.query.AddCell(SrcConstants.Relationships, nameof(SrcConstants.Relationships));


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
            cells.ShapeFixedCode = row[this.ShapeFixedCode];
            cells.ShapePermeablePlace = row[this.ShapePermeablePlace];
            cells.ShapePermeableX = row[this.ShapePermeableX];
            cells.ShapePermeableY = row[this.ShapePermeableY];
            cells.ShapePlaceFlip = row[this.ShapePlaceFlip];
            cells.ShapePlaceStyle = row[this.ShapePlaceStyle];
            cells.ShapePlowCode = row[this.ShapePlowCode];
            cells.ShapeRouteStyle = row[this.ShapeRouteStyle];
            cells.ShapeSplit = row[this.ShapeSplit];
            cells.ShapeSplittable = row[this.ShapeSplittable];
            cells.DisplayLevel = row[this.DisplayLevel];
            cells.Relationships = row[this.Relationships];
            return cells;
        }
    }
}