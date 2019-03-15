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
 
            public ShapeFormatCellsReader() : base(VisioAutomation.ShapeSheet.CellGroups.CellGroupReaderType.SingleRow)
            {
                InitializeQuery();
            }

            public override ShapeFormatCells ToCellGroup(ShapeSheet.Internal.ArraySegment<string> row)
            {

                var cells = new ShapeFormatCells();
                var cols = this.query_singlerow.Columns;

                string getcellvalue(string name)
                {
                    return row[cols[name].Ordinal];
                }

                cells.FillBackground = getcellvalue(nameof(ShapeFormatCells.FillBackground));
                cells.FillBackgroundTransparency= getcellvalue(nameof(ShapeFormatCells.FillBackgroundTransparency));
                cells.FillForeground = getcellvalue(nameof(ShapeFormatCells.FillForeground));
                cells.FillForegroundTransparency = getcellvalue(nameof(ShapeFormatCells.FillForegroundTransparency));
                cells.FillPattern = getcellvalue(nameof(ShapeFormatCells.FillPattern));
                cells.FillShadowObliqueAngle = getcellvalue(nameof(ShapeFormatCells.FillShadowObliqueAngle));
                cells.FillShadowOffsetX = getcellvalue(nameof(ShapeFormatCells.FillShadowOffsetX));
                cells.FillShadowOffsetY = getcellvalue(nameof(ShapeFormatCells.FillShadowOffsetY));
                cells.FillShadowScaleFactor = getcellvalue(nameof(ShapeFormatCells.FillShadowScaleFactor));
                cells.FillShadowType = getcellvalue(nameof(ShapeFormatCells.FillShadowType));
                cells.FillShadowBackground = getcellvalue(nameof(ShapeFormatCells.FillShadowBackground));
                cells.FillShadowBackgroundTransparency = getcellvalue(nameof(ShapeFormatCells.FillShadowBackgroundTransparency));
                cells.FillShadowForeground = getcellvalue(nameof(ShapeFormatCells.FillShadowForeground));
                cells.FillShadowForegroundTransparency = getcellvalue(nameof(ShapeFormatCells.FillShadowForegroundTransparency));
                cells.FillShadowPattern = getcellvalue(nameof(ShapeFormatCells.FillShadowPattern));
                cells.LineBeginArrow = getcellvalue(nameof(ShapeFormatCells.LineBeginArrow));
                cells.LineBeginArrowSize = getcellvalue(nameof(ShapeFormatCells.LineBeginArrowSize));
                cells.LineEndArrow = getcellvalue(nameof(ShapeFormatCells.LineEndArrow));
                cells.LineEndArrowSize = getcellvalue(nameof(ShapeFormatCells.LineEndArrowSize));
                cells.LineCap = getcellvalue(nameof(ShapeFormatCells.LineCap));
                cells.LineColor = getcellvalue(nameof(ShapeFormatCells.LineColor));
                cells.LineColorTransparency = getcellvalue(nameof(ShapeFormatCells.LineColorTransparency));
                cells.LinePattern = getcellvalue(nameof(ShapeFormatCells.LinePattern));
                cells.LineWeight = getcellvalue(nameof(ShapeFormatCells.LineWeight));
                cells.LineRounding = getcellvalue(nameof(ShapeFormatCells.LineRounding));
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

            public ShapeLayoutCellsReader() : base(VisioAutomation.ShapeSheet.CellGroups.CellGroupReaderType.SingleRow)
            {
                InitializeQuery();
            }

            public override ShapeLayoutCells ToCellGroup(ShapeSheet.Internal.ArraySegment<string> row)
            {
                var cells = new ShapeLayoutCells();
                var cols = this.query_singlerow.Columns;

                string getcellvalue(string name)
                {
                    return row[cols[name].Ordinal];
                }

                cells.ConnectorFixedCode = getcellvalue(nameof(ShapeLayoutCells.ConnectorFixedCode));
                cells.LineJumpCode = getcellvalue(nameof(ShapeLayoutCells.LineJumpCode));
                cells.LineJumpDirX = getcellvalue(nameof(ShapeLayoutCells.LineJumpDirX));
                cells.LineJumpDirY = getcellvalue(nameof(ShapeLayoutCells.LineJumpDirY));
                cells.LineJumpStyle = getcellvalue(nameof(ShapeLayoutCells.LineJumpStyle));
                cells.LineRouteExt = getcellvalue(nameof(ShapeLayoutCells.LineRouteExt));
                cells.ShapeFixedCode = getcellvalue(nameof(ShapeLayoutCells.ShapeFixedCode));
                cells.ShapePermeablePlace = getcellvalue(nameof(ShapeLayoutCells.ShapePermeablePlace));
                cells.ShapePermeableX = getcellvalue(nameof(ShapeLayoutCells.ShapePermeableX));
                cells.ShapePermeableY = getcellvalue(nameof(ShapeLayoutCells.ShapePermeableY));
                cells.ShapePlaceFlip = getcellvalue(nameof(ShapeLayoutCells.ShapePlaceFlip));
                cells.ShapePlaceStyle = getcellvalue(nameof(ShapeLayoutCells.ShapePlaceStyle));
                cells.ShapePlowCode = getcellvalue(nameof(ShapeLayoutCells.ShapePlowCode));
                cells.ShapeRouteStyle = getcellvalue(nameof(ShapeLayoutCells.ShapeRouteStyle));
                cells.ShapeSplit = getcellvalue(nameof(ShapeLayoutCells.ShapeSplit));
                cells.ShapeSplittable = getcellvalue(nameof(ShapeLayoutCells.ShapeSplittable));
                cells.ShapeDisplayLevel = getcellvalue(nameof(ShapeLayoutCells.ShapeDisplayLevel));
                cells.Relationships = getcellvalue(nameof(ShapeLayoutCells.Relationships));

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
            public ShapeXFormCellsReader() : base(VisioAutomation.ShapeSheet.CellGroups.CellGroupReaderType.SingleRow)
            {
                InitializeQuery();
            }

            public override ShapeXFormCells ToCellGroup(ShapeSheet.Internal.ArraySegment<string> row)
            {
                var cells = new ShapeXFormCells();

                var cols = this.query_singlerow.Columns;

                string getcellvalue(string name)
                {
                    return row[cols[name].Ordinal];
                }

                cells.PinX = getcellvalue(nameof(ShapeXFormCells.PinX));
                cells.PinY = getcellvalue(nameof(ShapeXFormCells.PinY));
                cells.LocPinX = getcellvalue(nameof(ShapeXFormCells.LocPinX));
                cells.LocPinY = getcellvalue(nameof(ShapeXFormCells.LocPinY));
                cells.Width = getcellvalue(nameof(ShapeXFormCells.Width));
                cells.Height = getcellvalue(nameof(ShapeXFormCells.Height));
                cells.Angle = getcellvalue(nameof(ShapeXFormCells.Angle));

                return cells;
            }
        }
    }
}