using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Shapes
{
    class ShapeFormatCellsReader : SingleRowReader<Shapes.ShapeFormatCells>
    {
        public CellColumn FillBkgnd { get; set; }
        public CellColumn FillBkgndTrans { get; set; }
        public CellColumn FillForegnd { get; set; }
        public CellColumn FillForegndTrans { get; set; }
        public CellColumn FillPattern { get; set; }
        public CellColumn ShapeShdwObliqueAngle { get; set; }
        public CellColumn ShapeShdwOffsetX { get; set; }
        public CellColumn ShapeShdwOffsetY { get; set; }
        public CellColumn ShapeShdwScaleFactor { get; set; }
        public CellColumn ShapeShdwType { get; set; }
        public CellColumn ShdwBkgnd { get; set; }
        public CellColumn ShdwBkgndTrans { get; set; }
        public CellColumn ShdwForegnd { get; set; }
        public CellColumn ShdwForegndTrans { get; set; }
        public CellColumn ShdwPattern { get; set; }
        public CellColumn BeginArrow { get; set; }
        public CellColumn BeginArrowSize { get; set; }
        public CellColumn EndArrow { get; set; }
        public CellColumn EndArrowSize { get; set; }
        public CellColumn LineColor { get; set; }
        public CellColumn LineCap { get; set; }
        public CellColumn LineColorTrans { get; set; }
        public CellColumn LinePattern { get; set; }
        public CellColumn LineWeight { get; set; }
        public CellColumn Rounding { get; set; }

        public ShapeFormatCellsReader()
        {
            
            this.FillBkgnd = this.query.AddCell(SrcConstants.FillBackground, nameof(SrcConstants.FillBackground));
            this.FillBkgndTrans = this.query.AddCell(SrcConstants.FillBackgroundTransparency, nameof(SrcConstants.FillBackgroundTransparency));
            this.FillForegnd = this.query.AddCell(SrcConstants.FillForeground, nameof(SrcConstants.FillForeground));
            this.FillForegndTrans = this.query.AddCell(SrcConstants.FillForegroundTransparency, nameof(SrcConstants.FillForegroundTransparency));
            this.FillPattern = this.query.AddCell(SrcConstants.FillPattern, nameof(SrcConstants.FillPattern));
            this.ShapeShdwObliqueAngle = this.query.AddCell(SrcConstants.FillShadowObliqueAngle, nameof(SrcConstants.FillShadowObliqueAngle));
            this.ShapeShdwOffsetX = this.query.AddCell(SrcConstants.FillShadowOffsetX, nameof(SrcConstants.FillShadowOffsetX));
            this.ShapeShdwOffsetY = this.query.AddCell(SrcConstants.FillShadowOffsetY, nameof(SrcConstants.FillShadowOffsetY));
            this.ShapeShdwScaleFactor = this.query.AddCell(SrcConstants.FillShadowScaleFactor, nameof(SrcConstants.FillShadowScaleFactor));
            this.ShapeShdwType = this.query.AddCell(SrcConstants.FillShadowType, nameof(SrcConstants.FillShadowType));
            this.ShdwBkgnd = this.query.AddCell(SrcConstants.FillShadowBackground, nameof(SrcConstants.FillShadowBackground));
            this.ShdwBkgndTrans = this.query.AddCell(SrcConstants.FillShadowBackgroundTransparency, nameof(SrcConstants.FillShadowBackgroundTransparency));
            this.ShdwForegnd = this.query.AddCell(SrcConstants.FillShadowForeground, nameof(SrcConstants.FillShadowForeground));
            this.ShdwForegndTrans = this.query.AddCell(SrcConstants.FillShadowForegroundTransparency, nameof(SrcConstants.FillShadowForegroundTransparency));
            this.ShdwPattern = this.query.AddCell(SrcConstants.FillShadowPattern, nameof(SrcConstants.FillShadowPattern));
            this.BeginArrow = this.query.AddCell(SrcConstants.LineBeginArrow, nameof(SrcConstants.LineBeginArrow));
            this.BeginArrowSize = this.query.AddCell(SrcConstants.LineBeginArrowSize, nameof(SrcConstants.LineBeginArrowSize));
            this.EndArrow = this.query.AddCell(SrcConstants.LineEndArrow, nameof(SrcConstants.LineEndArrow));
            this.EndArrowSize = this.query.AddCell(SrcConstants.LineEndArrowSize, nameof(SrcConstants.LineEndArrowSize));
            this.LineColor = this.query.AddCell(SrcConstants.LineColor, nameof(SrcConstants.LineColor));
            this.LineCap = this.query.AddCell(SrcConstants.LineCap, nameof(SrcConstants.LineCap));
            this.LineColorTrans = this.query.AddCell(SrcConstants.LineColorTransparency, nameof(SrcConstants.LineColorTransparency));
            this.LinePattern = this.query.AddCell(SrcConstants.LinePattern, nameof(SrcConstants.LinePattern));
            this.LineWeight = this.query.AddCell(SrcConstants.LineWeight, nameof(SrcConstants.LineWeight));
            this.Rounding = this.query.AddCell(SrcConstants.LineRounding, nameof(SrcConstants.LineRounding));
        }

        public override Shapes.ShapeFormatCells CellDataToCellGroup(VisioAutomation.Utilities.ArraySegment<ShapeSheet.CellData> row)
        {
            var cells = new Shapes.ShapeFormatCells();
            cells.FillBkgnd = row[this.FillBkgnd];
            cells.FillBkgndTrans = row[this.FillBkgndTrans];
            cells.FillForegnd = row[this.FillForegnd];
            cells.FillForegndTrans = row[this.FillForegndTrans];
            cells.FillPattern = row[this.FillPattern];
            cells.ShapeShdwObliqueAngle = row[this.ShapeShdwObliqueAngle];
            cells.ShapeShdwOffsetX = row[this.ShapeShdwOffsetX];
            cells.ShapeShdwOffsetY = row[this.ShapeShdwOffsetY];
            cells.ShapeShdwScaleFactor = row[this.ShapeShdwScaleFactor];
            cells.ShapeShdwType = row[this.ShapeShdwType];
            cells.ShdwBkgnd = row[this.ShdwBkgnd];
            cells.ShdwBkgndTrans = row[this.ShdwBkgndTrans];
            cells.ShdwForegnd = row[this.ShdwForegnd];
            cells.ShdwForegndTrans = row[this.ShdwForegndTrans];
            cells.ShdwPattern = row[this.ShdwPattern];
            cells.BeginArrow = row[this.BeginArrow];
            cells.BeginArrowSize = row[this.BeginArrowSize];
            cells.EndArrow = row[this.EndArrow];
            cells.EndArrowSize = row[this.EndArrowSize];
            cells.LineCap = row[this.LineCap];
            cells.LineColor = row[this.LineColor];
            cells.LineColorTrans = row[this.LineColorTrans];
            cells.LinePattern = row[this.LinePattern];
            cells.LineWeight = row[this.LineWeight];
            cells.Rounding = row[this.Rounding];
            return cells;
        }

    }
}