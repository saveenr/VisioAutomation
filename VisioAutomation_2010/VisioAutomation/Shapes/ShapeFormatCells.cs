using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Shapes
{
    public class ShapeFormatCells : CellGroupBase
    {
        public CellValueLiteral FillBackground { get; set; }
        public CellValueLiteral FillBackgroundTransparency { get; set; }
        public CellValueLiteral FillForeground { get; set; }
        public CellValueLiteral FillForegroundTransparency { get; set; }
        public CellValueLiteral FillPattern { get; set; }
        public CellValueLiteral FillShadowObliqueAngle { get; set; }
        public CellValueLiteral FillShadowOffsetX { get; set; }
        public CellValueLiteral FillShadowOffsetY { get; set; }
        public CellValueLiteral FillShadowScaleFactor { get; set; }
        public CellValueLiteral FillShadowType { get; set; }
        public CellValueLiteral FillShadowBackground { get; set; }
        public CellValueLiteral FillShadowBackgroundTransparency { get; set; }
        public CellValueLiteral FillShadowForeground { get; set; }
        public CellValueLiteral FillShadowForegroundTransparency { get; set; }
        public CellValueLiteral FillShadowPattern { get; set; }
        public CellValueLiteral LineBeginArrow { get; set; }
        public CellValueLiteral LineBeginArrowSize { get; set; }
        public CellValueLiteral LineEndArrow { get; set; }
        public CellValueLiteral LineEndArrowSize { get; set; }
        public CellValueLiteral LineCap { get; set; }
        public CellValueLiteral LineColor { get; set; }
        public CellValueLiteral LineColorTransparency { get; set; }
        public CellValueLiteral LinePattern { get; set; }
        public CellValueLiteral LineWeight { get; set; }
        public CellValueLiteral LineRounding { get; set; }

        public override IEnumerable<SrcValuePair> SrcValuePairs
        {
            get
            {
                yield return SrcValuePair.Create(SrcConstants.FillBackground, this.FillBackground);
                yield return SrcValuePair.Create(SrcConstants.FillBackgroundTransparency, this.FillBackgroundTransparency);
                yield return SrcValuePair.Create(SrcConstants.FillForeground, this.FillForeground);
                yield return SrcValuePair.Create(SrcConstants.FillForegroundTransparency, this.FillForegroundTransparency);
                yield return SrcValuePair.Create(SrcConstants.FillPattern, this.FillPattern);
                yield return SrcValuePair.Create(SrcConstants.FillShadowObliqueAngle, this.FillShadowObliqueAngle);
                yield return SrcValuePair.Create(SrcConstants.FillShadowOffsetX, this.FillShadowOffsetX);
                yield return SrcValuePair.Create(SrcConstants.FillShadowOffsetY, this.FillShadowOffsetY);
                yield return SrcValuePair.Create(SrcConstants.FillShadowScaleFactor, this.FillShadowScaleFactor);
                yield return SrcValuePair.Create(SrcConstants.FillShadowType, this.FillShadowType);
                yield return SrcValuePair.Create(SrcConstants.FillShadowBackground, this.FillShadowBackground);
                yield return SrcValuePair.Create(SrcConstants.FillShadowBackgroundTransparency, this.FillShadowBackgroundTransparency);
                yield return SrcValuePair.Create(SrcConstants.FillShadowForeground, this.FillShadowForeground);
                yield return SrcValuePair.Create(SrcConstants.FillShadowForegroundTransparency, this.FillShadowForegroundTransparency);
                yield return SrcValuePair.Create(SrcConstants.FillShadowPattern, this.FillShadowPattern);
                yield return SrcValuePair.Create(SrcConstants.LineBeginArrow, this.LineBeginArrow);
                yield return SrcValuePair.Create(SrcConstants.LineBeginArrowSize, this.LineBeginArrowSize);
                yield return SrcValuePair.Create(SrcConstants.LineEndArrow, this.LineEndArrow);
                yield return SrcValuePair.Create(SrcConstants.LineEndArrowSize, this.LineEndArrowSize);
                yield return SrcValuePair.Create(SrcConstants.LineCap, this.LineCap);
                yield return SrcValuePair.Create(SrcConstants.LineColor, this.LineColor);
                yield return SrcValuePair.Create(SrcConstants.LineColorTransparency, this.LineColorTransparency);
                yield return SrcValuePair.Create(SrcConstants.LinePattern, this.LinePattern);
                yield return SrcValuePair.Create(SrcConstants.LineWeight, this.LineWeight);
                yield return SrcValuePair.Create(SrcConstants.LineRounding, this.LineRounding);
            }
        }


        public static List<ShapeFormatCells> GetCells(IVisio.Page page, IList<int> shapeids, CellValueType type)
        {
            var reader = lazy_reader.Value;
            return reader.GetCellsSingleRow(page, shapeids, type);
        }

        public static ShapeFormatCells GetCells(IVisio.Shape shape, CellValueType type)
        {
            var reader = lazy_reader.Value;
            return reader.GetCellsSingleRow(shape, type);
        }

        private static readonly System.Lazy<ShapeFormatCellsReader> lazy_reader = new System.Lazy<ShapeFormatCellsReader>();

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

            public ShapeFormatCellsReader()
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

    }
}

