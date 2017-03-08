using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Shapes.Layout
{
    class ShapeLayoutCellsReader : SingleRowReader<Shapes.Layout.ShapeLayoutCells>
    {
        public CellColumn ConnectorFixedCode { get; set; }
        public CellColumn ConnectorLineJumpCode { get; set; }
        public CellColumn ConnectorLineJumpDirX { get; set; }
        public CellColumn ConnectorLineJumpDirY { get; set; }
        public CellColumn ConnectorLineJumpStyle { get; set; }
        public CellColumn ConnectorLineRouteExt { get; set; }
        public CellColumn FixedCode { get; set; }
        public CellColumn PermeablePlace { get; set; }
        public CellColumn PermeableX { get; set; }
        public CellColumn PermeableY { get; set; }
        public CellColumn PlaceFlip { get; set; }
        public CellColumn PlaceStyle { get; set; }
        public CellColumn PlowCode { get; set; }
        public CellColumn RouteStyle { get; set; }
        public CellColumn Split { get; set; }
        public CellColumn Splittable { get; set; }
        public CellColumn DisplayLevel { get; set; }
        public CellColumn Relationships { get; set; }

        public ShapeLayoutCellsReader() 
        {
            this.ConnectorFixedCode = this.query.AddCell(SrcConstants.ShapeLayoutConnectorFixedCode, nameof(SrcConstants.ShapeLayoutConnectorFixedCode));
            this.ConnectorLineJumpCode = this.query.AddCell(SrcConstants.ShapeLayoutConnectorLineJumpCode, nameof(SrcConstants.ShapeLayoutConnectorLineJumpCode));
            this.ConnectorLineJumpDirX = this.query.AddCell(SrcConstants.ShapeLayoutConnectorLineJumpDirX, nameof(SrcConstants.ShapeLayoutConnectorLineJumpDirX));
            this.ConnectorLineJumpDirY = this.query.AddCell(SrcConstants.ShapeLayoutConnectorLineJumpDirY, nameof(SrcConstants.ShapeLayoutConnectorLineJumpDirY));
            this.ConnectorLineJumpStyle = this.query.AddCell(SrcConstants.ShapeLayoutConnectorLineJumpStyle, nameof(SrcConstants.ShapeLayoutConnectorLineJumpStyle));
            this.ConnectorLineRouteExt = this.query.AddCell(SrcConstants.ShapeLayoutConnectorLineRouteExt, nameof(SrcConstants.ShapeLayoutConnectorLineRouteExt));
            this.FixedCode = this.query.AddCell(SrcConstants.ShapeLayoutFixedCode, nameof(SrcConstants.ShapeLayoutFixedCode));
            this.PermeablePlace = this.query.AddCell(SrcConstants.ShapeLayoutPermeablePlace, nameof(SrcConstants.ShapeLayoutPermeablePlace));
            this.PermeableX = this.query.AddCell(SrcConstants.ShapeLayoutPermeableX, nameof(SrcConstants.ShapeLayoutPermeableX));
            this.PermeableY = this.query.AddCell(SrcConstants.ShapeLayoutPermeableY, nameof(SrcConstants.ShapeLayoutPermeableY));
            this.PlaceFlip = this.query.AddCell(SrcConstants.ShapeLayoutPlaceFlip, nameof(SrcConstants.ShapeLayoutPlaceFlip));
            this.PlaceStyle = this.query.AddCell(SrcConstants.ShapeLayoutPlaceStyle, nameof(SrcConstants.ShapeLayoutPlaceStyle));
            this.PlowCode = this.query.AddCell(SrcConstants.ShapeLayoutPlowCode, nameof(SrcConstants.ShapeLayoutPlowCode));
            this.RouteStyle = this.query.AddCell(SrcConstants.ShapeLayoutRouteStyle, nameof(SrcConstants.ShapeLayoutRouteStyle));
            this.Split = this.query.AddCell(SrcConstants.ShapeLayoutSplit, nameof(SrcConstants.ShapeLayoutSplit));
            this.Splittable = this.query.AddCell(SrcConstants.ShapeLayoutSplittable, nameof(SrcConstants.ShapeLayoutSplittable));
            this.DisplayLevel = this.query.AddCell(SrcConstants.ShapeLayoutDisplayLevel, nameof(SrcConstants.ShapeLayoutDisplayLevel));
            this.Relationships = this.query.AddCell(SrcConstants.ShapeLayoutRelationships, nameof(SrcConstants.ShapeLayoutRelationships));
        }

        public override Shapes.Layout.ShapeLayoutCells CellDataToCellGroup(VisioAutomation.Utilities.ArraySegment<ShapeSheet.CellData> row)
        {
            var cells = new Shapes.Layout.ShapeLayoutCells();
            cells.ConnectorFixedCode = row[this.ConnectorFixedCode];
            cells.ConnectorLineJumpCode = row[this.ConnectorLineJumpCode];
            cells.ConnectorLineJumpDirX = row[this.ConnectorLineJumpDirX];
            cells.ConnectorLineJumpDirY = row[this.ConnectorLineJumpDirY];
            cells.ConnectorLineJumpStyle = row[this.ConnectorLineJumpStyle];
            cells.ConnectorLineRouteExt = row[this.ConnectorLineRouteExt];
            cells.FixedCode = row[this.FixedCode];
            cells.PermeablePlace = row[this.PermeablePlace];
            cells.PermeableX = row[this.PermeableX];
            cells.PermeableY = row[this.PermeableY];
            cells.PlaceFlip = row[this.PlaceFlip];
            cells.PlaceStyle = row[this.PlaceStyle];
            cells.PlowCode = row[this.PlowCode];
            cells.RouteStyle = row[this.RouteStyle];
            cells.Split = row[this.Split];
            cells.Splittable = row[this.Splittable];
            cells.DisplayLevel = row[this.DisplayLevel];
            cells.Relationships = row[this.Relationships];
            return cells;
        }
    }
}