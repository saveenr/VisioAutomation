using System.Collections.Generic;
using VisioAutomation.Extensions;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.Query;



namespace VisioAutomation.Shapes
{
    public static class ShapeHelper
    {
        /// <summary>
        /// Enumerates all shapes contained by a set of shapes recursively
        /// </summary>
        /// <param name="shapes">the set of shapes to start the enumeration</param>
        /// <returns>The enumeration</returns>
        public static List<IVisio.Shape> GetNestedShapes(IEnumerable<IVisio.Shape> shapes)
        {
            if (shapes == null)
            {
                throw new System.ArgumentNullException(nameof(shapes));
            }

            var result = new List<IVisio.Shape>();
            var stack = new Stack<IVisio.Shape>(shapes);

            while (stack.Count > 0)
            {
                var s = stack.Pop();
                var subshapes = s.Shapes;
                if (subshapes.Count > 0)
                {
                    foreach (var child in subshapes.ToEnumerable())
                    {
                        stack.Push(child);
                    }
                }

                result.Add(s);
            }

            return result;
        }

        public static List<IVisio.Shape> GetNestedShapes(IVisio.Shape shape)
        {
            if (shape== null)
            {
                throw new System.ArgumentNullException(nameof(shape));
            }

            var shapes = new[] {shape};

            return ShapeHelper.GetNestedShapes(shapes);
        }

        public static List<IVisio.Shape> GetShapesFromIDs(IVisio.Shapes shapes, IList<short> shapeids)
        {
            var shape_objs = new List<IVisio.Shape>(shapeids.Count);
            foreach (short shapeid in shapeids)
            {
                var shape = shapes.ItemFromID16[shapeid];
                shape_objs.Add(shape);
            }
            return shape_objs;
        }



        public static List<ShapeFormatCells> GetShapeFormatCells(IVisio.Page page, IList<int> shapeids, CellValueType type)
        {
            var reader = shape_format_lazy_reader.Value;
            return reader.GetCellsSingleRow(page, shapeids, type);
        }

        public static ShapeFormatCells GetShapeFormatCells(IVisio.Shape shape, CellValueType type)
        {
            var reader = shape_format_lazy_reader.Value;
            return reader.GetCellsSingleRow(shape, type);
        }

        private static readonly System.Lazy<ShapeFormatCellsReader> shape_format_lazy_reader = new System.Lazy<ShapeFormatCellsReader>();

        class ShapeFormatCellsReader : CellGroupReader<ShapeFormatCells>
        {
 
            public ShapeFormatCellsReader() : base(new VisioAutomation.ShapeSheet.Query.CellQuery())
            {

                var temp_cells = new ShapeFormatCells();
                foreach (var pair in temp_cells.NamedSrcValuePairs)
                {
                    this.query_singlerow.Columns.Add(pair.Src, pair.Name);
                }

            }

            public override ShapeFormatCells ToCellGroup(ShapeSheet.Internal.ArraySegment<string> row)
            {

                var cells = new ShapeFormatCells();
                var cols = this.query_singlerow.Columns;
                cells.FillBackground = row[cols[nameof(ShapeFormatCells.FillBackground)].Ordinal];
                cells.FillBackgroundTransparency= row[cols[nameof(ShapeFormatCells.FillBackgroundTransparency)].Ordinal];
                cells.FillForeground = row[cols[nameof(ShapeFormatCells.FillForeground)].Ordinal];
                cells.FillForegroundTransparency = row[cols[nameof(ShapeFormatCells.FillForegroundTransparency)].Ordinal];
                cells.FillPattern = row[cols[nameof(ShapeFormatCells.FillPattern)].Ordinal];
                cells.FillShadowObliqueAngle = row[cols[nameof(ShapeFormatCells.FillShadowObliqueAngle)].Ordinal];
                cells.FillShadowOffsetX = row[cols[nameof(ShapeFormatCells.FillShadowOffsetX)].Ordinal];
                cells.FillShadowOffsetY = row[cols[nameof(ShapeFormatCells.FillShadowOffsetY)].Ordinal];
                cells.FillShadowScaleFactor = row[cols[nameof(ShapeFormatCells.FillShadowScaleFactor)].Ordinal];
                cells.FillShadowType = row[cols[nameof(ShapeFormatCells.FillShadowType)].Ordinal];
                cells.FillShadowBackground = row[cols[nameof(ShapeFormatCells.FillShadowBackground)].Ordinal];
                cells.FillShadowBackgroundTransparency = row[cols[nameof(ShapeFormatCells.FillShadowBackgroundTransparency)].Ordinal];
                cells.FillShadowForeground = row[cols[nameof(ShapeFormatCells.FillShadowForeground)].Ordinal];
                cells.FillShadowForegroundTransparency = row[cols[nameof(ShapeFormatCells.FillShadowForegroundTransparency)].Ordinal];
                cells.FillShadowPattern = row[cols[nameof(ShapeFormatCells.FillShadowPattern)].Ordinal];
                cells.LineBeginArrow = row[cols[nameof(ShapeFormatCells.LineBeginArrow)].Ordinal];
                cells.LineBeginArrowSize = row[cols[nameof(ShapeFormatCells.LineBeginArrowSize)].Ordinal];
                cells.LineEndArrow = row[cols[nameof(ShapeFormatCells.LineEndArrow)].Ordinal];
                cells.LineEndArrowSize = row[cols[nameof(ShapeFormatCells.LineEndArrowSize)].Ordinal];
                cells.LineCap = row[cols[nameof(ShapeFormatCells.LineCap)].Ordinal];
                cells.LineColor = row[cols[nameof(ShapeFormatCells.LineColor)].Ordinal];
                cells.LineColorTransparency = row[cols[nameof(ShapeFormatCells.LineColorTransparency)].Ordinal];
                cells.LinePattern = row[cols[nameof(ShapeFormatCells.LinePattern)].Ordinal];
                cells.LineWeight = row[cols[nameof(ShapeFormatCells.LineWeight)].Ordinal];
                cells.LineRounding = row[cols[nameof(ShapeFormatCells.LineRounding)].Ordinal];
                return cells;
            }

        }


        public static List<ShapeLayoutCells> GetShapeLayoutCells(IVisio.Page page, IList<int> shapeids, CellValueType type)
        {
            var reader = ShapeLayoutCells_lazy_reader.Value;
            return reader.GetCellsSingleRow(page, shapeids, type);
        }

        public static ShapeLayoutCells GetShapeLayoutCells(IVisio.Shape shape, CellValueType type)
        {
            var reader = ShapeLayoutCells_lazy_reader.Value;
            return reader.GetCellsSingleRow(shape, type);
        }

        private static readonly System.Lazy<ShapeLayoutCellsReader> ShapeLayoutCells_lazy_reader = new System.Lazy<ShapeLayoutCellsReader>();

        class ShapeLayoutCellsReader : CellGroupReader<ShapeLayoutCells>
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

            public ShapeLayoutCellsReader() : base(new VisioAutomation.ShapeSheet.Query.CellQuery())
            {
                this.ConnectorFixedCode = this.query_singlerow.Columns.Add(SrcConstants.ShapeLayoutConnectorFixedCode, nameof(this.ConnectorFixedCode));
                this.LineJumpCode = this.query_singlerow.Columns.Add(SrcConstants.ShapeLayoutLineJumpCode, nameof(this.LineJumpCode));
                this.LineJumpDirX = this.query_singlerow.Columns.Add(SrcConstants.ShapeLayoutLineJumpDirX, nameof(this.LineJumpDirX));
                this.LineJumpDirY = this.query_singlerow.Columns.Add(SrcConstants.ShapeLayoutLineJumpDirY, nameof(this.LineJumpDirY));
                this.LineJumpStyle = this.query_singlerow.Columns.Add(SrcConstants.ShapeLayoutLineJumpStyle, nameof(this.LineJumpStyle));
                this.LineRouteExt = this.query_singlerow.Columns.Add(SrcConstants.ShapeLayoutLineRouteExt, nameof(this.LineRouteExt));
                this.ShapeFixedCode = this.query_singlerow.Columns.Add(SrcConstants.ShapeLayoutShapeFixedCode, nameof(this.ShapeFixedCode));
                this.ShapePermeablePlace = this.query_singlerow.Columns.Add(SrcConstants.ShapeLayoutShapePermeablePlace, nameof(this.ShapePermeablePlace));
                this.ShapePermeableX = this.query_singlerow.Columns.Add(SrcConstants.ShapeLayoutShapePermeableX, nameof(this.ShapePermeableX));
                this.ShapePermeableY = this.query_singlerow.Columns.Add(SrcConstants.ShapeLayoutShapePermeableY, nameof(this.ShapePermeableY));
                this.ShapePlaceFlip = this.query_singlerow.Columns.Add(SrcConstants.ShapeLayoutShapePlaceFlip, nameof(this.ShapePlaceFlip));
                this.ShapePlaceStyle = this.query_singlerow.Columns.Add(SrcConstants.ShapeLayoutShapePlaceStyle, nameof(this.ShapePlaceStyle));
                this.ShapePlowCode = this.query_singlerow.Columns.Add(SrcConstants.ShapeLayoutShapePlowCode, nameof(this.ShapePlowCode));
                this.ShapeRouteStyle = this.query_singlerow.Columns.Add(SrcConstants.ShapeLayoutShapeRouteStyle, nameof(this.ShapeRouteStyle));
                this.ShapeSplit = this.query_singlerow.Columns.Add(SrcConstants.ShapeLayoutShapeSplit, nameof(this.ShapeSplit));
                this.ShapeSplittable = this.query_singlerow.Columns.Add(SrcConstants.ShapeLayoutShapeSplittable, nameof(this.ShapeSplittable));
                this.ShapeDisplayLevel = this.query_singlerow.Columns.Add(SrcConstants.ShapeLayoutShapeDisplayLevel, nameof(this.ShapeDisplayLevel));
                this.Relationships = this.query_singlerow.Columns.Add(SrcConstants.ShapeLayoutRelationships, nameof(this.Relationships));
            }

            public override ShapeLayoutCells ToCellGroup(ShapeSheet.Internal.ArraySegment<string> row)
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


        public static List<ShapeXFormCells> GetShapeXFormCells(IVisio.Page page, IList<int> shapeids, CellValueType type)
        {
            var reader = ShapeXFormCells_lazy_reader.Value;
            return reader.GetCellsSingleRow(page, shapeids, type);
        }

        public static ShapeXFormCells GetShapeXFormCells(IVisio.Shape shape, CellValueType type)
        {
            var reader = ShapeXFormCells_lazy_reader.Value;
            return reader.GetCellsSingleRow(shape, type);
        }

        private static readonly System.Lazy<ShapeXFormCellsReader> ShapeXFormCells_lazy_reader = new System.Lazy<ShapeXFormCellsReader>();

        class ShapeXFormCellsReader : CellGroupReader<ShapeXFormCells>
        {
            public CellColumn Width { get; set; }
            public CellColumn Height { get; set; }
            public CellColumn PinX { get; set; }
            public CellColumn PinY { get; set; }
            public CellColumn LocPinX { get; set; }
            public CellColumn LocPinY { get; set; }
            public CellColumn Angle { get; set; }

            public ShapeXFormCellsReader() : base(new VisioAutomation.ShapeSheet.Query.CellQuery())
            {
                this.PinX = this.query_singlerow.Columns.Add(SrcConstants.XFormPinX, nameof(this.PinX));
                this.PinY = this.query_singlerow.Columns.Add(SrcConstants.XFormPinY, nameof(this.PinY));
                this.LocPinX = this.query_singlerow.Columns.Add(SrcConstants.XFormLocPinX, nameof(this.LocPinX));
                this.LocPinY = this.query_singlerow.Columns.Add(SrcConstants.XFormLocPinY, nameof(this.LocPinY));
                this.Width = this.query_singlerow.Columns.Add(SrcConstants.XFormWidth, nameof(this.Width));
                this.Height = this.query_singlerow.Columns.Add(SrcConstants.XFormHeight, nameof(this.Height));
                this.Angle = this.query_singlerow.Columns.Add(SrcConstants.XFormAngle, nameof(this.Angle));
            }

            public override ShapeXFormCells ToCellGroup(ShapeSheet.Internal.ArraySegment<string> row)
            {
                var cells = new ShapeXFormCells();
                cells.PinX = row[this.PinX];
                cells.PinY = row[this.PinY];
                cells.LocPinX = row[this.LocPinX];
                cells.LocPinY = row[this.LocPinY];
                cells.Width = row[this.Width];
                cells.Height = row[this.Height];
                cells.Angle = row[this.Angle];
                return cells;
            }
        }
    }
}