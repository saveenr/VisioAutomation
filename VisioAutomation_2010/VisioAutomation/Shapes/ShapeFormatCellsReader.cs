using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet.Queries;
using SRCCON = VisioAutomation.ShapeSheet.SRCConstants;

namespace VisioAutomation.Shapes
{
    class ShapeFormatCellsReader : SingleRowReader<Shapes.ShapeFormatCells>
    {
        public ColumnQuery FillBkgnd { get; set; }
        public ColumnQuery FillBkgndTrans { get; set; }
        public ColumnQuery FillForegnd { get; set; }
        public ColumnQuery FillForegndTrans { get; set; }
        public ColumnQuery FillPattern { get; set; }
        public ColumnQuery ShapeShdwObliqueAngle { get; set; }
        public ColumnQuery ShapeShdwOffsetX { get; set; }
        public ColumnQuery ShapeShdwOffsetY { get; set; }
        public ColumnQuery ShapeShdwScaleFactor { get; set; }
        public ColumnQuery ShapeShdwType { get; set; }
        public ColumnQuery ShdwBkgnd { get; set; }
        public ColumnQuery ShdwBkgndTrans { get; set; }
        public ColumnQuery ShdwForegnd { get; set; }
        public ColumnQuery ShdwForegndTrans { get; set; }
        public ColumnQuery ShdwPattern { get; set; }
        public ColumnQuery BeginArrow { get; set; }
        public ColumnQuery BeginArrowSize { get; set; }
        public ColumnQuery EndArrow { get; set; }
        public ColumnQuery EndArrowSize { get; set; }
        public ColumnQuery LineColor { get; set; }
        public ColumnQuery LineCap { get; set; }
        public ColumnQuery LineColorTrans { get; set; }
        public ColumnQuery LinePattern { get; set; }
        public ColumnQuery LineWeight { get; set; }
        public ColumnQuery Rounding { get; set; }

        public ShapeFormatCellsReader()
        {
            
            this.FillBkgnd = this.query.AddCell(SRCCON.FillBkgnd, nameof(SRCCON.FillBkgnd));
            this.FillBkgndTrans = this.query.AddCell(SRCCON.FillBkgndTrans, nameof(SRCCON.FillBkgndTrans));
            this.FillForegnd = this.query.AddCell(SRCCON.FillForegnd, nameof(SRCCON.FillForegnd));
            this.FillForegndTrans = this.query.AddCell(SRCCON.FillForegndTrans, nameof(SRCCON.FillForegndTrans));
            this.FillPattern = this.query.AddCell(SRCCON.FillPattern, nameof(SRCCON.FillPattern));
            this.ShapeShdwObliqueAngle = this.query.AddCell(SRCCON.ShapeShdwObliqueAngle, nameof(SRCCON.ShapeShdwObliqueAngle));
            this.ShapeShdwOffsetX = this.query.AddCell(SRCCON.ShapeShdwOffsetX, nameof(SRCCON.ShapeShdwOffsetX));
            this.ShapeShdwOffsetY = this.query.AddCell(SRCCON.ShapeShdwOffsetY, nameof(SRCCON.ShapeShdwOffsetY));
            this.ShapeShdwScaleFactor = this.query.AddCell(SRCCON.ShapeShdwScaleFactor, nameof(SRCCON.ShapeShdwScaleFactor));
            this.ShapeShdwType = this.query.AddCell(SRCCON.ShapeShdwType, nameof(SRCCON.ShapeShdwType));
            this.ShdwBkgnd = this.query.AddCell(SRCCON.ShdwBkgnd, nameof(SRCCON.ShdwBkgnd));
            this.ShdwBkgndTrans = this.query.AddCell(SRCCON.ShdwBkgndTrans, nameof(SRCCON.ShdwBkgndTrans));
            this.ShdwForegnd = this.query.AddCell(SRCCON.ShdwForegnd, nameof(SRCCON.ShdwForegnd));
            this.ShdwForegndTrans = this.query.AddCell(SRCCON.ShdwForegndTrans, nameof(SRCCON.ShdwForegndTrans));
            this.ShdwPattern = this.query.AddCell(SRCCON.ShdwPattern, nameof(SRCCON.ShdwPattern));
            this.BeginArrow = this.query.AddCell(SRCCON.BeginArrow, nameof(SRCCON.BeginArrow));
            this.BeginArrowSize = this.query.AddCell(SRCCON.BeginArrowSize, nameof(SRCCON.BeginArrowSize));
            this.EndArrow = this.query.AddCell(SRCCON.EndArrow, nameof(SRCCON.EndArrow));
            this.EndArrowSize = this.query.AddCell(SRCCON.EndArrowSize, nameof(SRCCON.EndArrowSize));
            this.LineColor = this.query.AddCell(SRCCON.LineColor, nameof(SRCCON.LineColor));
            this.LineCap = this.query.AddCell(SRCCON.LineCap, nameof(SRCCON.LineCap));
            this.LineColorTrans = this.query.AddCell(SRCCON.LineColorTrans, nameof(SRCCON.LineColorTrans));
            this.LinePattern = this.query.AddCell(SRCCON.LinePattern, nameof(SRCCON.LinePattern));
            this.LineWeight = this.query.AddCell(SRCCON.LineWeight, nameof(SRCCON.LineWeight));
            this.Rounding = this.query.AddCell(SRCCON.Rounding, nameof(SRCCON.Rounding));
        }

        public override Shapes.ShapeFormatCells CellDataToCellGroup(ShapeSheet.CellData[] row)
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