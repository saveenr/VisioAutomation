using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet.Query;
using SRCCON = VisioAutomation.ShapeSheet.SRCConstants;

namespace VisioAutomation.Shapes
{
    class ShapeFormatCellsReader : SingleRowReader<Shapes.ShapeFormatCells>
    {
        public ColumnCell FillBkgnd { get; set; }
        public ColumnCell FillBkgndTrans { get; set; }
        public ColumnCell FillForegnd { get; set; }
        public ColumnCell FillForegndTrans { get; set; }
        public ColumnCell FillPattern { get; set; }
        public ColumnCell ShapeShdwObliqueAngle { get; set; }
        public ColumnCell ShapeShdwOffsetX { get; set; }
        public ColumnCell ShapeShdwOffsetY { get; set; }
        public ColumnCell ShapeShdwScaleFactor { get; set; }
        public ColumnCell ShapeShdwType { get; set; }
        public ColumnCell ShdwBkgnd { get; set; }
        public ColumnCell ShdwBkgndTrans { get; set; }
        public ColumnCell ShdwForegnd { get; set; }
        public ColumnCell ShdwForegndTrans { get; set; }
        public ColumnCell ShdwPattern { get; set; }
        public ColumnCell BeginArrow { get; set; }
        public ColumnCell BeginArrowSize { get; set; }
        public ColumnCell EndArrow { get; set; }
        public ColumnCell EndArrowSize { get; set; }
        public ColumnCell LineColor { get; set; }
        public ColumnCell LineCap { get; set; }
        public ColumnCell LineColorTrans { get; set; }
        public ColumnCell LinePattern { get; set; }
        public ColumnCell LineWeight { get; set; }
        public ColumnCell Rounding { get; set; }

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