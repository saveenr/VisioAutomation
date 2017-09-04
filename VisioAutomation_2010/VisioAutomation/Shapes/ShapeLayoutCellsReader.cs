using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Shapes
{
    class ShapeLayoutCellsReader : ReaderSingleRow<ShapeLayoutCells>
    {
        public CellColumn ConnectorFixedCode { get; set; }
        public CellColumn LineJumpCode { get; set; }
        public CellColumn LineJumpDirX { get; set; }
        public CellColumn LineJumpDirY { get; set; }
        public CellColumn LineJumpStyle { get; set; }
        public CellColumn LineRouteExt { get; set; }
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
        public CellColumn ShapeDisplayLevel { get; set; }
        public CellColumn Relationships { get; set; }

        public ShapeLayoutCellsReader() 
        {
            this.ConnectorFixedCode = this.query.AddCell(SrcConstants.ShapeLayoutConnectorFixedCode, nameof(SrcConstants.ShapeLayoutConnectorFixedCode));
            this.LineJumpCode = this.query.AddCell(SrcConstants.ShapeLayoutLineJumpCode, nameof(SrcConstants.ShapeLayoutLineJumpCode));
            this.LineJumpDirX = this.query.AddCell(SrcConstants.ShapeLayoutLineJumpDirX, nameof(SrcConstants.ShapeLayoutLineJumpDirX));
            this.LineJumpDirY = this.query.AddCell(SrcConstants.ShapeLayoutLineJumpDirY, nameof(SrcConstants.ShapeLayoutLineJumpDirY));
            this.LineJumpStyle = this.query.AddCell(SrcConstants.ShapeLayoutLineJumpStyle, nameof(SrcConstants.ShapeLayoutLineJumpStyle));
            this.LineRouteExt = this.query.AddCell(SrcConstants.ShapeLayoutLineRouteExt, nameof(SrcConstants.ShapeLayoutLineRouteExt));
            this.ShapeFixedCode = this.query.AddCell(SrcConstants.ShapeLayoutShapeFixedCode, nameof(SrcConstants.ShapeLayoutShapeFixedCode));
            this.ShapePermeablePlace = this.query.AddCell(SrcConstants.ShapeLayoutShapePermeablePlace, nameof(SrcConstants.ShapeLayoutShapePermeablePlace));
            this.ShapePermeableX = this.query.AddCell(SrcConstants.ShapeLayoutShapePermeableX, nameof(SrcConstants.ShapeLayoutShapePermeableX));
            this.ShapePermeableY = this.query.AddCell(SrcConstants.ShapeLayoutShapePermeableY, nameof(SrcConstants.ShapeLayoutShapePermeableY));
            this.ShapePlaceFlip = this.query.AddCell(SrcConstants.ShapeLayoutShapePlaceFlip, nameof(SrcConstants.ShapeLayoutShapePlaceFlip));
            this.ShapePlaceStyle = this.query.AddCell(SrcConstants.ShapeLayoutShapePlaceStyle, nameof(SrcConstants.ShapeLayoutShapePlaceStyle));
            this.ShapePlowCode = this.query.AddCell(SrcConstants.ShapeLayoutShapePlowCode, nameof(SrcConstants.ShapeLayoutShapePlowCode));
            this.ShapeRouteStyle = this.query.AddCell(SrcConstants.ShapeLayoutShapeRouteStyle, nameof(SrcConstants.ShapeLayoutShapeRouteStyle));
            this.ShapeSplit = this.query.AddCell(SrcConstants.ShapeLayoutShapeSplit, nameof(SrcConstants.ShapeLayoutShapeSplit));
            this.ShapeSplittable = this.query.AddCell(SrcConstants.ShapeLayoutShapeSplittable, nameof(SrcConstants.ShapeLayoutShapeSplittable));
            this.ShapeDisplayLevel = this.query.AddCell(SrcConstants.ShapeLayoutShapeDisplayLevel, nameof(SrcConstants.ShapeLayoutShapeDisplayLevel));
            this.Relationships = this.query.AddCell(SrcConstants.ShapeLayoutRelationships, nameof(SrcConstants.ShapeLayoutRelationships));
        }

        public override ShapeLayoutCells CellDataToCellGroup(VisioAutomation.Utilities.ArraySegment<ShapeSheet.CellData> row)
        {
            var cells = new ShapeLayoutCells();
            cells.ConnectorFixedCode = row[this.ConnectorFixedCode];
            cells.LineJumpCode = row[this.LineJumpCode];
            cells.LineJumpDirX = row[this.LineJumpDirX];
            cells.LineJumpDirY = row[this.LineJumpDirY];
            cells.LineJumpStyle = row[this.LineJumpStyle];
            cells.LineRouteExt = row[this.LineRouteExt];
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
            cells.ShapeDisplayLevel = row[this.ShapeDisplayLevel];
            cells.Relationships = row[this.Relationships];
            return cells;
        }
    }
}