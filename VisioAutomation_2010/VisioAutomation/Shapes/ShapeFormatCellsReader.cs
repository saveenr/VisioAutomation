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
            
            this.FillBackground = this.query.AddCell(SrcConstants.FillBackground, nameof(SrcConstants.FillBackground));
            this.FillBackgroundTransparency = this.query.AddCell(SrcConstants.FillBackgroundTransparency, nameof(SrcConstants.FillBackgroundTransparency));
            this.FillForeground = this.query.AddCell(SrcConstants.FillForeground, nameof(SrcConstants.FillForeground));
            this.FillForegroundTransparency = this.query.AddCell(SrcConstants.FillForegroundTransparency, nameof(SrcConstants.FillForegroundTransparency));
            this.FillPattern = this.query.AddCell(SrcConstants.FillPattern, nameof(SrcConstants.FillPattern));
            this.FillShadowObliqueAngle = this.query.AddCell(SrcConstants.FillShadowObliqueAngle, nameof(SrcConstants.FillShadowObliqueAngle));
            this.FillShadowOffsetX = this.query.AddCell(SrcConstants.FillShadowOffsetX, nameof(SrcConstants.FillShadowOffsetX));
            this.FillShadowOffsetY = this.query.AddCell(SrcConstants.FillShadowOffsetY, nameof(SrcConstants.FillShadowOffsetY));
            this.FillShadowScaleFactor = this.query.AddCell(SrcConstants.FillShadowScaleFactor, nameof(SrcConstants.FillShadowScaleFactor));
            this.FillShadowType = this.query.AddCell(SrcConstants.FillShadowType, nameof(SrcConstants.FillShadowType));
            this.FillShadowBackground = this.query.AddCell(SrcConstants.FillShadowBackground, nameof(SrcConstants.FillShadowBackground));
            this.FillShadowBackgroundTransparency = this.query.AddCell(SrcConstants.FillShadowBackgroundTransparency, nameof(SrcConstants.FillShadowBackgroundTransparency));
            this.FillShadowForeground = this.query.AddCell(SrcConstants.FillShadowForeground, nameof(SrcConstants.FillShadowForeground));
            this.FillShadowForegroundTransparency = this.query.AddCell(SrcConstants.FillShadowForegroundTransparency, nameof(SrcConstants.FillShadowForegroundTransparency));
            this.FillShadowPattern = this.query.AddCell(SrcConstants.FillShadowPattern, nameof(SrcConstants.FillShadowPattern));
            this.LineBeginArrow = this.query.AddCell(SrcConstants.LineBeginArrow, nameof(SrcConstants.LineBeginArrow));
            this.LineBeginArrowSize = this.query.AddCell(SrcConstants.LineBeginArrowSize, nameof(SrcConstants.LineBeginArrowSize));
            this.LineEndArrow = this.query.AddCell(SrcConstants.LineEndArrow, nameof(SrcConstants.LineEndArrow));
            this.LineEndArrowSize = this.query.AddCell(SrcConstants.LineEndArrowSize, nameof(SrcConstants.LineEndArrowSize));
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