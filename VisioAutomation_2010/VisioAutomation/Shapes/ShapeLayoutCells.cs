using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using IVisio= Microsoft.Office.Interop.Visio;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Shapes
{
    public class ShapeLayoutCells : ShapeSheet.CellGroups.CellGroupSingleRow
    {
        public VisioAutomation.ShapeSheet.CellValueLiteral ConnectorFixedCode { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral LineJumpCode { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral LineJumpDirX { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral LineJumpDirY { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral LineJumpStyle { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral LineRouteExt { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral ShapeFixedCode { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral ShapePermeablePlace { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral ShapePermeableX { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral ShapePermeableY { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral ShapePlaceFlip { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral ShapePlaceStyle { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral ShapePlowCode { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral ShapeRouteStyle { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral ShapeSplit { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral ShapeSplittable { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral ShapeDisplayLevel { get; set; } // new in visio 2010
        public VisioAutomation.ShapeSheet.CellValueLiteral Relationships { get; set; } // new in visio 2010

        public override IEnumerable<SrcValuePair> SrcValuePairs
        {
            get
            {
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ShapeLayoutConnectorFixedCode, this.ConnectorFixedCode.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ShapeLayoutLineJumpCode, this.LineJumpCode.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ShapeLayoutLineJumpDirX, this.LineJumpDirX.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ShapeLayoutLineJumpDirY, this.LineJumpDirY.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ShapeLayoutLineJumpStyle, this.LineJumpStyle.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ShapeLayoutLineRouteExt, this.LineRouteExt.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ShapeLayoutShapeFixedCode, this.ShapeFixedCode.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ShapeLayoutShapePermeablePlace, this.ShapePermeablePlace.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ShapeLayoutShapePermeableX, this.ShapePermeableX.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ShapeLayoutShapePermeableY, this.ShapePermeableY.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ShapeLayoutShapePlaceFlip, this.ShapePlaceFlip.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ShapeLayoutShapePlaceStyle, this.ShapePlaceStyle.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ShapeLayoutShapePlowCode, this.ShapePlowCode.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ShapeLayoutShapeRouteStyle, this.ShapeRouteStyle.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ShapeLayoutShapeSplit, this.ShapeSplit.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ShapeLayoutShapeSplittable, this.ShapeSplittable.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ShapeLayoutShapeDisplayLevel, this.ShapeDisplayLevel.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.ShapeLayoutRelationships, this.Relationships.Value);
            }
        }
        
        public static List<ShapeLayoutCells> GetValues(IVisio.Page page, IList<int> shapeids, CellValueType cvt)
        {
            var query = ShapeLayoutCells.lazy_query.Value;
            return query.GetValues(page, shapeids, cvt);
        }

        public static ShapeLayoutCells GetValues(IVisio.Shape shape, CellValueType cvt)
        {
            var query = ShapeLayoutCells.lazy_query.Value;
            return query.GetValues(shape, cvt);
        }

        private static readonly System.Lazy<ShapeLayoutCellsReader> lazy_query = new System.Lazy<ShapeLayoutCellsReader>();

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

            public override ShapeLayoutCells CellDataToCellGroup(VisioAutomation.Utilities.ArraySegment<string> row)
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
}