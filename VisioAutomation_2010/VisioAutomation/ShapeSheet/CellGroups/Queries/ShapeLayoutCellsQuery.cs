using VisioAutomation.ShapeSheet.Queries.Columns;
using SRCCON = VisioAutomation.ShapeSheet.SRCConstants;
using VisioAutomation.Extensions;

namespace VisioAutomation.ShapeSheet.CellGroups.Queries
{
    class ShapeLayoutCellsQuery : CellGroupSingleRowQuery<Shapes.Layout.ShapeLayoutCells>
    {
        public ColumnQuery ConFixedCode { get; set; }
        public ColumnQuery ConLineJumpCode { get; set; }
        public ColumnQuery ConLineJumpDirX { get; set; }
        public ColumnQuery ConLineJumpDirY { get; set; }
        public ColumnQuery ConLineJumpStyle { get; set; }
        public ColumnQuery ConLineRouteExt { get; set; }
        public ColumnQuery ShapeFixedCode { get; set; }
        public ColumnQuery ShapePermeablePlace { get; set; }
        public ColumnQuery ShapePermeableX { get; set; }
        public ColumnQuery ShapePermeableY { get; set; }
        public ColumnQuery ShapePlaceFlip { get; set; }
        public ColumnQuery ShapePlaceStyle { get; set; }
        public ColumnQuery ShapePlowCode { get; set; }
        public ColumnQuery ShapeRouteStyle { get; set; }
        public ColumnQuery ShapeSplit { get; set; }
        public ColumnQuery ShapeSplittable { get; set; }
        public ColumnQuery DisplayLevel { get; set; }
        public ColumnQuery Relationships { get; set; }

        public ShapeLayoutCellsQuery() 
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

        public override Shapes.Layout.ShapeLayoutCells CellDataToCellGroup(ShapeSheet.CellData[] row)
        {
            var cells = new Shapes.Layout.ShapeLayoutCells();
            cells.ConFixedCode = row[this.ConFixedCode].ToInt();
            cells.ConLineJumpCode = row[this.ConLineJumpCode].ToInt();
            cells.ConLineJumpDirX = row[this.ConLineJumpDirX].ToInt();
            cells.ConLineJumpDirY = row[this.ConLineJumpDirY].ToInt();
            cells.ConLineJumpStyle = row[this.ConLineJumpStyle].ToInt();
            cells.ConLineRouteExt = row[this.ConLineRouteExt].ToInt();
            cells.ShapeFixedCode = row[this.ShapeFixedCode].ToInt();
            cells.ShapePermeablePlace = row[this.ShapePermeablePlace].ToInt();
            cells.ShapePermeableX = row[this.ShapePermeableX].ToInt();
            cells.ShapePermeableY = row[this.ShapePermeableY].ToInt();
            cells.ShapePlaceFlip = row[this.ShapePlaceFlip].ToInt();
            cells.ShapePlaceStyle = row[this.ShapePlaceStyle].ToInt();
            cells.ShapePlowCode = row[this.ShapePlowCode].ToInt();
            cells.ShapeRouteStyle = row[this.ShapeRouteStyle].ToInt();
            cells.ShapeSplit = row[this.ShapeSplit].ToInt();
            cells.ShapeSplittable = row[this.ShapeSplittable].ToInt();
            cells.DisplayLevel = row[this.DisplayLevel].ToInt();
            cells.Relationships = row[this.Relationships].ToInt();
            return cells;
        }
    }
}