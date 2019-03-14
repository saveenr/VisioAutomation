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
            public CellColumn FillBackground { get; set; }
            public CellColumn FillBackgroundTransparency { get; set; }
            public CellColumn FillForeground { get; set; }
            public CellColumn FillForegroundTransparency { get; set; }
            public CellColumn FillPattern { get; set; }
            public CellColumn FillShadowObliqueAngle { get; set; }
            public CellColumn FillShadowOffsetX { get; set; }
            public CellColumn FillShadowOffsetY { get; set; }
            public CellColumn FillShadowScaleFactor { get; set; }
            public CellColumn FillShadowType { get; set; }
            public CellColumn FillShadowBackground { get; set; }
            public CellColumn FillShadowBackgroundTransparency { get; set; }
            public CellColumn FillShadowForeground { get; set; }
            public CellColumn FillShadowForegroundTransparency { get; set; }
            public CellColumn FillShadowPattern { get; set; }
            public CellColumn LineBeginArrow { get; set; }
            public CellColumn LineBeginArrowSize { get; set; }
            public CellColumn LineEndArrow { get; set; }
            public CellColumn LineEndArrowSize { get; set; }
            public CellColumn LineColor { get; set; }
            public CellColumn LineCap { get; set; }
            public CellColumn LineColorTransparency { get; set; }
            public CellColumn LinePattern { get; set; }
            public CellColumn LineWeight { get; set; }
            public CellColumn LineRounding { get; set; }

            public ShapeFormatCellsReader() : base(new VisioAutomation.ShapeSheet.Query.CellQuery())
            {

                this.FillBackground = this.query_singlerow.Columns.Add(SrcConstants.FillBackground, nameof(this.FillBackground));
                this.FillBackgroundTransparency = this.query_singlerow.Columns.Add(SrcConstants.FillBackgroundTransparency, nameof(this.FillBackgroundTransparency));
                this.FillForeground = this.query_singlerow.Columns.Add(SrcConstants.FillForeground, nameof(this.FillForeground));
                this.FillForegroundTransparency = this.query_singlerow.Columns.Add(SrcConstants.FillForegroundTransparency, nameof(this.FillForegroundTransparency));
                this.FillPattern = this.query_singlerow.Columns.Add(SrcConstants.FillPattern, nameof(this.FillPattern));
                this.FillShadowObliqueAngle = this.query_singlerow.Columns.Add(SrcConstants.FillShadowObliqueAngle, nameof(this.FillShadowObliqueAngle));
                this.FillShadowOffsetX = this.query_singlerow.Columns.Add(SrcConstants.FillShadowOffsetX, nameof(this.FillShadowOffsetX));
                this.FillShadowOffsetY = this.query_singlerow.Columns.Add(SrcConstants.FillShadowOffsetY, nameof(this.FillShadowOffsetY));
                this.FillShadowScaleFactor = this.query_singlerow.Columns.Add(SrcConstants.FillShadowScaleFactor, nameof(this.FillShadowScaleFactor));
                this.FillShadowType = this.query_singlerow.Columns.Add(SrcConstants.FillShadowType, nameof(this.FillShadowType));
                this.FillShadowBackground = this.query_singlerow.Columns.Add(SrcConstants.FillShadowBackground, nameof(this.FillShadowBackground));
                this.FillShadowBackgroundTransparency = this.query_singlerow.Columns.Add(SrcConstants.FillShadowBackgroundTransparency, nameof(this.FillShadowBackgroundTransparency));
                this.FillShadowForeground = this.query_singlerow.Columns.Add(SrcConstants.FillShadowForeground, nameof(this.FillShadowForeground));
                this.FillShadowForegroundTransparency = this.query_singlerow.Columns.Add(SrcConstants.FillShadowForegroundTransparency, nameof(this.FillShadowForegroundTransparency));
                this.FillShadowPattern = this.query_singlerow.Columns.Add(SrcConstants.FillShadowPattern, nameof(this.FillShadowPattern));
                this.LineBeginArrow = this.query_singlerow.Columns.Add(SrcConstants.LineBeginArrow, nameof(this.LineBeginArrow));
                this.LineBeginArrowSize = this.query_singlerow.Columns.Add(SrcConstants.LineBeginArrowSize, nameof(this.LineBeginArrowSize));
                this.LineEndArrow = this.query_singlerow.Columns.Add(SrcConstants.LineEndArrow, nameof(this.LineEndArrow));
                this.LineEndArrowSize = this.query_singlerow.Columns.Add(SrcConstants.LineEndArrowSize, nameof(this.LineEndArrowSize));
                this.LineColor = this.query_singlerow.Columns.Add(SrcConstants.LineColor, nameof(this.LineColor));
                this.LineCap = this.query_singlerow.Columns.Add(SrcConstants.LineCap, nameof(this.LineCap));
                this.LineColorTransparency = this.query_singlerow.Columns.Add(SrcConstants.LineColorTransparency, nameof(this.LineColorTransparency));
                this.LinePattern = this.query_singlerow.Columns.Add(SrcConstants.LinePattern, nameof(this.LinePattern));
                this.LineWeight = this.query_singlerow.Columns.Add(SrcConstants.LineWeight, nameof(this.LineWeight));
                this.LineRounding = this.query_singlerow.Columns.Add(SrcConstants.LineRounding, nameof(this.LineRounding));
            }

            public override ShapeFormatCells ToCellGroup(ShapeSheet.Internal.ArraySegment<string> row)
            {
                var cells = new ShapeFormatCells();
                cells.FillBackground = row[this.FillBackground];
                cells.FillBackgroundTransparency = row[this.FillBackgroundTransparency];
                cells.FillForeground = row[this.FillForeground];
                cells.FillForegroundTransparency = row[this.FillForegroundTransparency];
                cells.FillPattern = row[this.FillPattern];
                cells.FillShadowObliqueAngle = row[this.FillShadowObliqueAngle];
                cells.FillShadowOffsetX = row[this.FillShadowOffsetX];
                cells.FillShadowOffsetY = row[this.FillShadowOffsetY];
                cells.FillShadowScaleFactor = row[this.FillShadowScaleFactor];
                cells.FillShadowType = row[this.FillShadowType];
                cells.FillShadowBackground = row[this.FillShadowBackground];
                cells.FillShadowBackgroundTransparency = row[this.FillShadowBackgroundTransparency];
                cells.FillShadowForeground = row[this.FillShadowForeground];
                cells.FillShadowForegroundTransparency = row[this.FillShadowForegroundTransparency];
                cells.FillShadowPattern = row[this.FillShadowPattern];
                cells.LineBeginArrow = row[this.LineBeginArrow];
                cells.LineBeginArrowSize = row[this.LineBeginArrowSize];
                cells.LineEndArrow = row[this.LineEndArrow];
                cells.LineEndArrowSize = row[this.LineEndArrowSize];
                cells.LineCap = row[this.LineCap];
                cells.LineColor = row[this.LineColor];
                cells.LineColorTransparency = row[this.LineColorTransparency];
                cells.LinePattern = row[this.LinePattern];
                cells.LineWeight = row[this.LineWeight];
                cells.LineRounding = row[this.LineRounding];
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