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
            
            this.FillBkgnd = this.query.AddCell(SrcConstants.FillBkgnd, nameof(SrcConstants.FillBkgnd));
            this.FillBkgndTrans = this.query.AddCell(SrcConstants.FillBkgndTrans, nameof(SrcConstants.FillBkgndTrans));
            this.FillForegnd = this.query.AddCell(SrcConstants.FillForegnd, nameof(SrcConstants.FillForegnd));
            this.FillForegndTrans = this.query.AddCell(SrcConstants.FillForegndTrans, nameof(SrcConstants.FillForegndTrans));
            this.FillPattern = this.query.AddCell(SrcConstants.FillPattern, nameof(SrcConstants.FillPattern));
            this.ShapeShdwObliqueAngle = this.query.AddCell(SrcConstants.ShapeShdwObliqueAngle, nameof(SrcConstants.ShapeShdwObliqueAngle));
            this.ShapeShdwOffsetX = this.query.AddCell(SrcConstants.ShapeShdwOffsetX, nameof(SrcConstants.ShapeShdwOffsetX));
            this.ShapeShdwOffsetY = this.query.AddCell(SrcConstants.ShapeShdwOffsetY, nameof(SrcConstants.ShapeShdwOffsetY));
            this.ShapeShdwScaleFactor = this.query.AddCell(SrcConstants.ShapeShdwScaleFactor, nameof(SrcConstants.ShapeShdwScaleFactor));
            this.ShapeShdwType = this.query.AddCell(SrcConstants.ShapeShdwType, nameof(SrcConstants.ShapeShdwType));
            this.ShdwBkgnd = this.query.AddCell(SrcConstants.ShdwBkgnd, nameof(SrcConstants.ShdwBkgnd));
            this.ShdwBkgndTrans = this.query.AddCell(SrcConstants.ShdwBkgndTrans, nameof(SrcConstants.ShdwBkgndTrans));
            this.ShdwForegnd = this.query.AddCell(SrcConstants.ShdwForegnd, nameof(SrcConstants.ShdwForegnd));
            this.ShdwForegndTrans = this.query.AddCell(SrcConstants.ShdwForegndTrans, nameof(SrcConstants.ShdwForegndTrans));
            this.ShdwPattern = this.query.AddCell(SrcConstants.ShdwPattern, nameof(SrcConstants.ShdwPattern));
            this.BeginArrow = this.query.AddCell(SrcConstants.BeginArrow, nameof(SrcConstants.BeginArrow));
            this.BeginArrowSize = this.query.AddCell(SrcConstants.BeginArrowSize, nameof(SrcConstants.BeginArrowSize));
            this.EndArrow = this.query.AddCell(SrcConstants.EndArrow, nameof(SrcConstants.EndArrow));
            this.EndArrowSize = this.query.AddCell(SrcConstants.EndArrowSize, nameof(SrcConstants.EndArrowSize));
            this.LineColor = this.query.AddCell(SrcConstants.LineColor, nameof(SrcConstants.LineColor));
            this.LineCap = this.query.AddCell(SrcConstants.LineCap, nameof(SrcConstants.LineCap));
            this.LineColorTrans = this.query.AddCell(SrcConstants.LineColorTrans, nameof(SrcConstants.LineColorTrans));
            this.LinePattern = this.query.AddCell(SrcConstants.LinePattern, nameof(SrcConstants.LinePattern));
            this.LineWeight = this.query.AddCell(SrcConstants.LineWeight, nameof(SrcConstants.LineWeight));
            this.Rounding = this.query.AddCell(SrcConstants.Rounding, nameof(SrcConstants.Rounding));
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