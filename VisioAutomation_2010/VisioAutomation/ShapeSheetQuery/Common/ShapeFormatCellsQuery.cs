using SRCCON = VisioAutomation.ShapeSheet.SRCConstants;

namespace VisioAutomation.ShapeSheetQuery.Common
{
    class ShapeFormatCellsQuery : CellQuery
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

        public ShapeFormatCellsQuery()
        {
            
            this.FillBkgnd = this.AddCell(SRCCON.FillBkgnd, nameof(SRCCON.FillBkgnd));
            this.FillBkgndTrans = this.AddCell(SRCCON.FillBkgndTrans, nameof(SRCCON.FillBkgndTrans));
            this.FillForegnd = this.AddCell(SRCCON.FillForegnd, nameof(SRCCON.FillForegnd));
            this.FillForegndTrans = this.AddCell(SRCCON.FillForegndTrans, nameof(SRCCON.FillForegndTrans));
            this.FillPattern = this.AddCell(SRCCON.FillPattern, nameof(SRCCON.FillPattern));
            this.ShapeShdwObliqueAngle = this.AddCell(SRCCON.ShapeShdwObliqueAngle, nameof(SRCCON.ShapeShdwObliqueAngle));
            this.ShapeShdwOffsetX = this.AddCell(SRCCON.ShapeShdwOffsetX, nameof(SRCCON.ShapeShdwOffsetX));
            this.ShapeShdwOffsetY = this.AddCell(SRCCON.ShapeShdwOffsetY, nameof(SRCCON.ShapeShdwOffsetY));
            this.ShapeShdwScaleFactor = this.AddCell(SRCCON.ShapeShdwScaleFactor, nameof(SRCCON.ShapeShdwScaleFactor));
            this.ShapeShdwType = this.AddCell(SRCCON.ShapeShdwType, nameof(SRCCON.ShapeShdwType));
            this.ShdwBkgnd = this.AddCell(SRCCON.ShdwBkgnd, nameof(SRCCON.ShdwBkgnd));
            this.ShdwBkgndTrans = this.AddCell(SRCCON.ShdwBkgndTrans, nameof(SRCCON.ShdwBkgndTrans));
            this.ShdwForegnd = this.AddCell(SRCCON.ShdwForegnd, nameof(SRCCON.ShdwForegnd));
            this.ShdwForegndTrans = this.AddCell(SRCCON.ShdwForegndTrans, nameof(SRCCON.ShdwForegndTrans));
            this.ShdwPattern = this.AddCell(SRCCON.ShdwPattern, nameof(SRCCON.ShdwPattern));
            this.BeginArrow = this.AddCell(SRCCON.BeginArrow, nameof(SRCCON.BeginArrow));
            this.BeginArrowSize = this.AddCell(SRCCON.BeginArrowSize, nameof(SRCCON.BeginArrowSize));
            this.EndArrow = this.AddCell(SRCCON.EndArrow, nameof(SRCCON.EndArrow));
            this.EndArrowSize = this.AddCell(SRCCON.EndArrowSize, nameof(SRCCON.EndArrowSize));
            this.LineColor = this.AddCell(SRCCON.LineColor, nameof(SRCCON.LineColor));
            this.LineCap = this.AddCell(SRCCON.LineCap, nameof(SRCCON.LineCap));
            this.LineColorTrans = this.AddCell(SRCCON.LineColorTrans, nameof(SRCCON.LineColorTrans));
            this.LinePattern = this.AddCell(SRCCON.LinePattern, nameof(SRCCON.LinePattern));
            this.LineWeight = this.AddCell(SRCCON.LineWeight, nameof(SRCCON.LineWeight));
            this.Rounding = this.AddCell(SRCCON.Rounding, nameof(SRCCON.Rounding));


        }

        public Shapes.FormatCells GetCells(SectionResultRow<ShapeSheet.CellData<double>> row)
        {
            var cells = new Shapes.FormatCells();
            cells.FillBkgnd = Extensions.CellDataMethods.ToInt(row.Cells[this.FillBkgnd]);
            cells.FillBkgndTrans = row.Cells[this.FillBkgndTrans];
            cells.FillForegnd = Extensions.CellDataMethods.ToInt(row.Cells[this.FillForegnd]);
            cells.FillForegndTrans = row.Cells[this.FillForegndTrans];
            cells.FillPattern = Extensions.CellDataMethods.ToInt(row.Cells[this.FillPattern]);
            cells.ShapeShdwObliqueAngle = row.Cells[this.ShapeShdwObliqueAngle];
            cells.ShapeShdwOffsetX = row.Cells[this.ShapeShdwOffsetX];
            cells.ShapeShdwOffsetY = row.Cells[this.ShapeShdwOffsetY];
            cells.ShapeShdwScaleFactor = row.Cells[this.ShapeShdwScaleFactor];
            cells.ShapeShdwType = Extensions.CellDataMethods.ToInt(row.Cells[this.ShapeShdwType]);
            cells.ShdwBkgnd = Extensions.CellDataMethods.ToInt(row.Cells[this.ShdwBkgnd]);
            cells.ShdwBkgndTrans = row.Cells[this.ShdwBkgndTrans];
            cells.ShdwForegnd = Extensions.CellDataMethods.ToInt(row.Cells[this.ShdwForegnd]);
            cells.ShdwForegndTrans = row.Cells[this.ShdwForegndTrans];
            cells.ShdwPattern = Extensions.CellDataMethods.ToInt(row.Cells[this.ShdwPattern]);
            cells.BeginArrow = Extensions.CellDataMethods.ToInt(row.Cells[this.BeginArrow]);
            cells.BeginArrowSize = row.Cells[this.BeginArrowSize];
            cells.EndArrow = Extensions.CellDataMethods.ToInt(row.Cells[this.EndArrow]);
            cells.EndArrowSize = row.Cells[this.EndArrowSize];
            cells.LineCap = Extensions.CellDataMethods.ToInt(row.Cells[this.LineCap]);
            cells.LineColor = Extensions.CellDataMethods.ToInt(row.Cells[this.LineColor]);
            cells.LineColorTrans = row.Cells[this.LineColorTrans];
            cells.LinePattern = Extensions.CellDataMethods.ToInt(row.Cells[this.LinePattern]);
            cells.LineWeight = row.Cells[this.LineWeight];
            cells.Rounding = row.Cells[this.Rounding];
            return cells;
        }

    }
}