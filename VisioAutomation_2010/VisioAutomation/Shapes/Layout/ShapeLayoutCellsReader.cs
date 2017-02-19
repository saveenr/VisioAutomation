using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet.Query;
using SRCCON = VisioAutomation.ShapeSheet.SRCConstants;

namespace VisioAutomation.Shapes.Layout
{
    class ShapeLayoutCellsReader : SingleRowReader<Shapes.Layout.ShapeLayoutCells>
    {
        public ColumnCell ConFixedCode { get; set; }
        public ColumnCell ConLineJumpCode { get; set; }
        public ColumnCell ConLineJumpDirX { get; set; }
        public ColumnCell ConLineJumpDirY { get; set; }
        public ColumnCell ConLineJumpStyle { get; set; }
        public ColumnCell ConLineRouteExt { get; set; }
        public ColumnCell ShapeFixedCode { get; set; }
        public ColumnCell ShapePermeablePlace { get; set; }
        public ColumnCell ShapePermeableX { get; set; }
        public ColumnCell ShapePermeableY { get; set; }
        public ColumnCell ShapePlaceFlip { get; set; }
        public ColumnCell ShapePlaceStyle { get; set; }
        public ColumnCell ShapePlowCode { get; set; }
        public ColumnCell ShapeRouteStyle { get; set; }
        public ColumnCell ShapeSplit { get; set; }
        public ColumnCell ShapeSplittable { get; set; }
        public ColumnCell DisplayLevel { get; set; }
        public ColumnCell Relationships { get; set; }

        public ShapeLayoutCellsReader() 
        {
            this.ConFixedCode = this.query.AddCell(SRCCON.ConFixedCode, nameof(SRCCON.ConFixedCode));
            this.ConLineJumpCode = this.query.AddCell(SRCCON.ConLineJumpCode, nameof(SRCCON.ConLineJumpCode));
            this.ConLineJumpDirX = this.query.AddCell(SRCCON.ConLineJumpDirX, nameof(SRCCON.ConLineJumpDirX));
            this.ConLineJumpDirY = this.query.AddCell(SRCCON.ConLineJumpDirY, nameof(SRCCON.ConLineJumpDirY));
            this.ConLineJumpStyle = this.query.AddCell(SRCCON.ConLineJumpStyle, nameof(SRCCON.ConLineJumpStyle));
            this.ConLineRouteExt = this.query.AddCell(SRCCON.ConLineRouteExt, nameof(SRCCON.ConLineRouteExt));
            this.ShapeFixedCode = this.query.AddCell(SRCCON.ShapeFixedCode, nameof(SRCCON.ShapeFixedCode));
            this.ShapePermeablePlace = this.query.AddCell(SRCCON.ShapePermeablePlace, nameof(SRCCON.ShapePermeablePlace));
            this.ShapePermeableX = this.query.AddCell(SRCCON.ShapePermeableX, nameof(SRCCON.ShapePermeableX));
            this.ShapePermeableY = this.query.AddCell(SRCCON.ShapePermeableY, nameof(SRCCON.ShapePermeableY));
            this.ShapePlaceFlip = this.query.AddCell(SRCCON.ShapePlaceFlip, nameof(SRCCON.ShapePlaceFlip));
            this.ShapePlaceStyle = this.query.AddCell(SRCCON.ShapePlaceStyle, nameof(SRCCON.ShapePlaceStyle));
            this.ShapePlowCode = this.query.AddCell(SRCCON.ShapePlowCode, nameof(SRCCON.ShapePlowCode));
            this.ShapeRouteStyle = this.query.AddCell(SRCCON.ShapeRouteStyle, nameof(SRCCON.ShapeRouteStyle));
            this.ShapeSplit = this.query.AddCell(SRCCON.ShapeSplit, nameof(SRCCON.ShapeSplit));
            this.ShapeSplittable = this.query.AddCell(SRCCON.ShapeSplittable, nameof(SRCCON.ShapeSplittable));
            this.DisplayLevel = this.query.AddCell(SRCCON.DisplayLevel, nameof(SRCCON.DisplayLevel));
            this.Relationships = this.query.AddCell(SRCCON.Relationships, nameof(SRCCON.Relationships));


        }

        public override Shapes.Layout.ShapeLayoutCells CellDataToCellGroup(CellRange<ShapeSheet.CellData> row)
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