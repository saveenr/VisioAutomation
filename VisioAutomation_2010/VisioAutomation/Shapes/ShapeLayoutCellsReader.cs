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
            this.ConnectorFixedCode = this.query.Columns.Add(SrcConstants.ShapeLayoutConnectorFixedCode, nameof(SrcConstants.ShapeLayoutConnectorFixedCode));
            this.LineJumpCode = this.query.Columns.Add(SrcConstants.ShapeLayoutLineJumpCode, nameof(SrcConstants.ShapeLayoutLineJumpCode));
            this.LineJumpDirX = this.query.Columns.Add(SrcConstants.ShapeLayoutLineJumpDirX, nameof(SrcConstants.ShapeLayoutLineJumpDirX));
            this.LineJumpDirY = this.query.Columns.Add(SrcConstants.ShapeLayoutLineJumpDirY, nameof(SrcConstants.ShapeLayoutLineJumpDirY));
            this.LineJumpStyle = this.query.Columns.Add(SrcConstants.ShapeLayoutLineJumpStyle, nameof(SrcConstants.ShapeLayoutLineJumpStyle));
            this.LineRouteExt = this.query.Columns.Add(SrcConstants.ShapeLayoutLineRouteExt, nameof(SrcConstants.ShapeLayoutLineRouteExt));
            this.ShapeFixedCode = this.query.Columns.Add(SrcConstants.ShapeLayoutShapeFixedCode, nameof(SrcConstants.ShapeLayoutShapeFixedCode));
            this.ShapePermeablePlace = this.query.Columns.Add(SrcConstants.ShapeLayoutShapePermeablePlace, nameof(SrcConstants.ShapeLayoutShapePermeablePlace));
            this.ShapePermeableX = this.query.Columns.Add(SrcConstants.ShapeLayoutShapePermeableX, nameof(SrcConstants.ShapeLayoutShapePermeableX));
            this.ShapePermeableY = this.query.Columns.Add(SrcConstants.ShapeLayoutShapePermeableY, nameof(SrcConstants.ShapeLayoutShapePermeableY));
            this.ShapePlaceFlip = this.query.Columns.Add(SrcConstants.ShapeLayoutShapePlaceFlip, nameof(SrcConstants.ShapeLayoutShapePlaceFlip));
            this.ShapePlaceStyle = this.query.Columns.Add(SrcConstants.ShapeLayoutShapePlaceStyle, nameof(SrcConstants.ShapeLayoutShapePlaceStyle));
            this.ShapePlowCode = this.query.Columns.Add(SrcConstants.ShapeLayoutShapePlowCode, nameof(SrcConstants.ShapeLayoutShapePlowCode));
            this.ShapeRouteStyle = this.query.Columns.Add(SrcConstants.ShapeLayoutShapeRouteStyle, nameof(SrcConstants.ShapeLayoutShapeRouteStyle));
            this.ShapeSplit = this.query.Columns.Add(SrcConstants.ShapeLayoutShapeSplit, nameof(SrcConstants.ShapeLayoutShapeSplit));
            this.ShapeSplittable = this.query.Columns.Add(SrcConstants.ShapeLayoutShapeSplittable, nameof(SrcConstants.ShapeLayoutShapeSplittable));
            this.ShapeDisplayLevel = this.query.Columns.Add(SrcConstants.ShapeLayoutShapeDisplayLevel, nameof(SrcConstants.ShapeLayoutShapeDisplayLevel));
            this.Relationships = this.query.Columns.Add(SrcConstants.ShapeLayoutRelationships, nameof(SrcConstants.ShapeLayoutRelationships));
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