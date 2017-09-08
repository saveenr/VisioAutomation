using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Shapes
{
    public class ShapeFormatCells : ShapeSheet.CellGroups.CellGroupSingleRow
    {
        public VisioAutomation.ShapeSheet.CellValueLiteral FillBackground { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral FillBackgroundTransparency { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral FillForeground { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral FillForegroundTransparency { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral FillPattern { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral FillShadowObliqueAngle { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral FillShadowOffsetX { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral FillShadowOffsetY { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral FillShadowScaleFactor { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral FillShadowType { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral FillShadowBackground { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral FillShadowBackgroundTransparency { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral FillShadowForeground { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral FillShadowForegroundTransparency { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral FillShadowPattern { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral LineBeginArrow { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral LineBeginArrowSize { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral LineEndArrow { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral LineEndArrowSize { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral LineCap { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral LineColor { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral LineColorTransparency { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral LinePattern { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral LineWeight { get; set; }
        public VisioAutomation.ShapeSheet.CellValueLiteral LineRounding { get; set; }

        public override IEnumerable<SrcValuePair> SrcValuePairs
        {
            get
            {
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.FillBackground, this.FillBackground.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.FillBackgroundTransparency, this.FillBackgroundTransparency.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.FillForeground, this.FillForeground.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.FillForegroundTransparency, this.FillForegroundTransparency.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.FillPattern, this.FillPattern.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.FillShadowObliqueAngle, this.FillShadowObliqueAngle.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.FillShadowOffsetX, this.FillShadowOffsetX.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.FillShadowOffsetY, this.FillShadowOffsetY.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.FillShadowScaleFactor, this.FillShadowScaleFactor.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.FillShadowType, this.FillShadowType.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.FillShadowBackground, this.FillShadowBackground.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.FillShadowBackgroundTransparency, this.FillShadowBackgroundTransparency.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.FillShadowForeground, this.FillShadowForeground.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.FillShadowForegroundTransparency, this.FillShadowForegroundTransparency.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.FillShadowPattern, this.FillShadowPattern.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.LineBeginArrow, this.LineBeginArrow.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.LineBeginArrowSize, this.LineBeginArrowSize.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.LineEndArrow, this.LineEndArrow.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.LineEndArrowSize, this.LineEndArrowSize.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.LineCap, this.LineCap.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.LineColor, this.LineColor.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.LineColorTransparency, this.LineColorTransparency.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.LinePattern, this.LinePattern.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.LineWeight, this.LineWeight.Value);
                yield return SrcValuePair.Create(ShapeSheet.SrcConstants.LineRounding, this.LineRounding.Value);
            }
        }


        public static List<ShapeFormatCells> GetValues(IVisio.Page page, IList<int> shapeids, CellValueType cvt)
        {
            var query = ShapeFormatCells.lazy_query.Value;
            return query.GetValues(page, shapeids, cvt);
        }

        public static ShapeFormatCells GetValues(IVisio.Shape shape, CellValueType cvt)
        {
            var query = ShapeFormatCells.lazy_query.Value;
            return query.GetValues(shape, cvt);
        }

        private static readonly System.Lazy<ShapeFormatCellsReader> lazy_query = new System.Lazy<ShapeFormatCellsReader>();

        class ShapeFormatCellsReader : ReaderSingleRow<Shapes.ShapeFormatCells>
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
            public CellColumn LineColorTrans { get; set; }
            public CellColumn LinePattern { get; set; }
            public CellColumn LineWeight { get; set; }
            public CellColumn LineRounding { get; set; }

            public ShapeFormatCellsReader()
            {

                this.FillBackground = this.query.Columns.Add(SrcConstants.FillBackground, nameof(SrcConstants.FillBackground));
                this.FillBackgroundTransparency = this.query.Columns.Add(SrcConstants.FillBackgroundTransparency, nameof(SrcConstants.FillBackgroundTransparency));
                this.FillForeground = this.query.Columns.Add(SrcConstants.FillForeground, nameof(SrcConstants.FillForeground));
                this.FillForegroundTransparency = this.query.Columns.Add(SrcConstants.FillForegroundTransparency, nameof(SrcConstants.FillForegroundTransparency));
                this.FillPattern = this.query.Columns.Add(SrcConstants.FillPattern, nameof(SrcConstants.FillPattern));
                this.FillShadowObliqueAngle = this.query.Columns.Add(SrcConstants.FillShadowObliqueAngle, nameof(SrcConstants.FillShadowObliqueAngle));
                this.FillShadowOffsetX = this.query.Columns.Add(SrcConstants.FillShadowOffsetX, nameof(SrcConstants.FillShadowOffsetX));
                this.FillShadowOffsetY = this.query.Columns.Add(SrcConstants.FillShadowOffsetY, nameof(SrcConstants.FillShadowOffsetY));
                this.FillShadowScaleFactor = this.query.Columns.Add(SrcConstants.FillShadowScaleFactor, nameof(SrcConstants.FillShadowScaleFactor));
                this.FillShadowType = this.query.Columns.Add(SrcConstants.FillShadowType, nameof(SrcConstants.FillShadowType));
                this.FillShadowBackground = this.query.Columns.Add(SrcConstants.FillShadowBackground, nameof(SrcConstants.FillShadowBackground));
                this.FillShadowBackgroundTransparency = this.query.Columns.Add(SrcConstants.FillShadowBackgroundTransparency, nameof(SrcConstants.FillShadowBackgroundTransparency));
                this.FillShadowForeground = this.query.Columns.Add(SrcConstants.FillShadowForeground, nameof(SrcConstants.FillShadowForeground));
                this.FillShadowForegroundTransparency = this.query.Columns.Add(SrcConstants.FillShadowForegroundTransparency, nameof(SrcConstants.FillShadowForegroundTransparency));
                this.FillShadowPattern = this.query.Columns.Add(SrcConstants.FillShadowPattern, nameof(SrcConstants.FillShadowPattern));
                this.LineBeginArrow = this.query.Columns.Add(SrcConstants.LineBeginArrow, nameof(SrcConstants.LineBeginArrow));
                this.LineBeginArrowSize = this.query.Columns.Add(SrcConstants.LineBeginArrowSize, nameof(SrcConstants.LineBeginArrowSize));
                this.LineEndArrow = this.query.Columns.Add(SrcConstants.LineEndArrow, nameof(SrcConstants.LineEndArrow));
                this.LineEndArrowSize = this.query.Columns.Add(SrcConstants.LineEndArrowSize, nameof(SrcConstants.LineEndArrowSize));
                this.LineColor = this.query.Columns.Add(SrcConstants.LineColor, nameof(SrcConstants.LineColor));
                this.LineCap = this.query.Columns.Add(SrcConstants.LineCap, nameof(SrcConstants.LineCap));
                this.LineColorTrans = this.query.Columns.Add(SrcConstants.LineColorTransparency, nameof(SrcConstants.LineColorTransparency));
                this.LinePattern = this.query.Columns.Add(SrcConstants.LinePattern, nameof(SrcConstants.LinePattern));
                this.LineWeight = this.query.Columns.Add(SrcConstants.LineWeight, nameof(SrcConstants.LineWeight));
                this.LineRounding = this.query.Columns.Add(SrcConstants.LineRounding, nameof(SrcConstants.LineRounding));
            }

            public override Shapes.ShapeFormatCells CellDataToCellGroup(VisioAutomation.Utilities.ArraySegment<string> row)
            {
                var cells = new Shapes.ShapeFormatCells();
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
                cells.LineColorTransparency = row[this.LineColorTrans];
                cells.LinePattern = row[this.LinePattern];
                cells.LineWeight = row[this.LineWeight];
                cells.LineRounding = row[this.LineRounding];
                return cells;
            }

        }

    }
}

