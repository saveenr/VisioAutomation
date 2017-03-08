using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Shapes
{
    class ShapeFormatCellsReader : SingleRowReader<Shapes.ShapeFormatCells>
    {
        public CellColumn FillBackground { get; set; }
        public CellColumn FillBackgroundTransparency { get; set; }
        public CellColumn FillForeground { get; set; }
        public CellColumn FillForegroundTransparency { get; set; }
        public CellColumn FillPattern { get; set; }
        public CellColumn ShapeShadowObliqueAngle { get; set; }
        public CellColumn ShapeShadowOffsetX { get; set; }
        public CellColumn ShapeShadowOffsetY { get; set; }
        public CellColumn ShapeShadowScaleFactor { get; set; }
        public CellColumn ShapeShadowType { get; set; }
        public CellColumn ShadowBackground { get; set; }
        public CellColumn ShadowBackgroundTransparency { get; set; }
        public CellColumn ShadowForeground { get; set; }
        public CellColumn ShadowForegroundTransparency { get; set; }
        public CellColumn ShadowPattern { get; set; }
        public CellColumn BeginArrow { get; set; }
        public CellColumn BeginArrowSize { get; set; }
        public CellColumn EndArrow { get; set; }
        public CellColumn EndArrowSize { get; set; }
        public CellColumn LineColor { get; set; }
        public CellColumn LineCap { get; set; }
        public CellColumn LineColorTrans { get; set; }
        public CellColumn LinePattern { get; set; }
        public CellColumn LineWeight { get; set; }
        public CellColumn LineRounding { get; set; }

        public ShapeFormatCellsReader()
        {
            
            this.FillBackground = this.query.AddCell(SrcConstants.FillBackground, nameof(SrcConstants.FillBackground));
            this.FillBackgroundTransparency = this.query.AddCell(SrcConstants.FillBackgroundTransparency, nameof(SrcConstants.FillBackgroundTransparency));
            this.FillForeground = this.query.AddCell(SrcConstants.FillForeground, nameof(SrcConstants.FillForeground));
            this.FillForegroundTransparency = this.query.AddCell(SrcConstants.FillForegroundTransparency, nameof(SrcConstants.FillForegroundTransparency));
            this.FillPattern = this.query.AddCell(SrcConstants.FillPattern, nameof(SrcConstants.FillPattern));
            this.ShapeShadowObliqueAngle = this.query.AddCell(SrcConstants.FillShadowObliqueAngle, nameof(SrcConstants.FillShadowObliqueAngle));
            this.ShapeShadowOffsetX = this.query.AddCell(SrcConstants.FillShadowOffsetX, nameof(SrcConstants.FillShadowOffsetX));
            this.ShapeShadowOffsetY = this.query.AddCell(SrcConstants.FillShadowOffsetY, nameof(SrcConstants.FillShadowOffsetY));
            this.ShapeShadowScaleFactor = this.query.AddCell(SrcConstants.FillShadowScaleFactor, nameof(SrcConstants.FillShadowScaleFactor));
            this.ShapeShadowType = this.query.AddCell(SrcConstants.FillShadowType, nameof(SrcConstants.FillShadowType));
            this.ShadowBackground = this.query.AddCell(SrcConstants.FillShadowBackground, nameof(SrcConstants.FillShadowBackground));
            this.ShadowBackgroundTransparency = this.query.AddCell(SrcConstants.FillShadowBackgroundTransparency, nameof(SrcConstants.FillShadowBackgroundTransparency));
            this.ShadowForeground = this.query.AddCell(SrcConstants.FillShadowForeground, nameof(SrcConstants.FillShadowForeground));
            this.ShadowForegroundTransparency = this.query.AddCell(SrcConstants.FillShadowForegroundTransparency, nameof(SrcConstants.FillShadowForegroundTransparency));
            this.ShadowPattern = this.query.AddCell(SrcConstants.FillShadowPattern, nameof(SrcConstants.FillShadowPattern));
            this.BeginArrow = this.query.AddCell(SrcConstants.LineBeginArrow, nameof(SrcConstants.LineBeginArrow));
            this.BeginArrowSize = this.query.AddCell(SrcConstants.LineBeginArrowSize, nameof(SrcConstants.LineBeginArrowSize));
            this.EndArrow = this.query.AddCell(SrcConstants.LineEndArrow, nameof(SrcConstants.LineEndArrow));
            this.EndArrowSize = this.query.AddCell(SrcConstants.LineEndArrowSize, nameof(SrcConstants.LineEndArrowSize));
            this.LineColor = this.query.AddCell(SrcConstants.LineColor, nameof(SrcConstants.LineColor));
            this.LineCap = this.query.AddCell(SrcConstants.LineCap, nameof(SrcConstants.LineCap));
            this.LineColorTrans = this.query.AddCell(SrcConstants.LineColorTransparency, nameof(SrcConstants.LineColorTransparency));
            this.LinePattern = this.query.AddCell(SrcConstants.LinePattern, nameof(SrcConstants.LinePattern));
            this.LineWeight = this.query.AddCell(SrcConstants.LineWeight, nameof(SrcConstants.LineWeight));
            this.LineRounding = this.query.AddCell(SrcConstants.LineRounding, nameof(SrcConstants.LineRounding));
        }

        public override Shapes.ShapeFormatCells CellDataToCellGroup(VisioAutomation.Utilities.ArraySegment<ShapeSheet.CellData> row)
        {
            var cells = new Shapes.ShapeFormatCells();
            cells.FillBackground = row[this.FillBackground];
            cells.FillBackgroundTransparency = row[this.FillBackgroundTransparency];
            cells.FillForeground = row[this.FillForeground];
            cells.FillForegroundTransparency = row[this.FillForegroundTransparency];
            cells.FillPattern = row[this.FillPattern];
            cells.FillShadowObliqueAngle = row[this.ShapeShadowObliqueAngle];
            cells.FillShadowOffsetX = row[this.ShapeShadowOffsetX];
            cells.FillShadowOffsetY = row[this.ShapeShadowOffsetY];
            cells.FillShadowScaleFactor = row[this.ShapeShadowScaleFactor];
            cells.FillShadowType = row[this.ShapeShadowType];
            cells.FillShadowBackground = row[this.ShadowBackground];
            cells.FillShadowBackgroundTransparency = row[this.ShadowBackgroundTransparency];
            cells.FillShadowForeground = row[this.ShadowForeground];
            cells.FillShadowForegroundTransparency = row[this.ShadowForegroundTransparency];
            cells.FillShadowPattern = row[this.ShadowPattern];
            cells.LineBeginArrow = row[this.BeginArrow];
            cells.LineBeginArrowSize = row[this.BeginArrowSize];
            cells.LineEndArrow = row[this.EndArrow];
            cells.LineEndArrowSize = row[this.EndArrowSize];
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