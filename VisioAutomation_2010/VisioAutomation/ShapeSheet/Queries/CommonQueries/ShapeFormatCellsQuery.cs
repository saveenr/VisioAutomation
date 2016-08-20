using VisioAutomation.ShapeSheet.Queries.Columns;
using SRCCON = VisioAutomation.ShapeSheet.SRCConstants;

namespace VisioAutomation.ShapeSheet.Queries.CommonQueries
{
    class ShapeFormatCellsQuery : Query
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

        public Shapes.FormatCells GetCells(ShapeSheet.CellData<double>[] row)
        {
            var cells = new Shapes.FormatCells();
            cells.FillBkgnd = Extensions.CellDataMethods.ToInt(row[this.FillBkgnd]);
            cells.FillBkgndTrans = row[this.FillBkgndTrans];
            cells.FillForegnd = Extensions.CellDataMethods.ToInt(row[this.FillForegnd]);
            cells.FillForegndTrans = row[this.FillForegndTrans];
            cells.FillPattern = Extensions.CellDataMethods.ToInt(row[this.FillPattern]);
            cells.ShapeShdwObliqueAngle = row[this.ShapeShdwObliqueAngle];
            cells.ShapeShdwOffsetX = row[this.ShapeShdwOffsetX];
            cells.ShapeShdwOffsetY = row[this.ShapeShdwOffsetY];
            cells.ShapeShdwScaleFactor = row[this.ShapeShdwScaleFactor];
            cells.ShapeShdwType = Extensions.CellDataMethods.ToInt(row[this.ShapeShdwType]);
            cells.ShdwBkgnd = Extensions.CellDataMethods.ToInt(row[this.ShdwBkgnd]);
            cells.ShdwBkgndTrans = row[this.ShdwBkgndTrans];
            cells.ShdwForegnd = Extensions.CellDataMethods.ToInt(row[this.ShdwForegnd]);
            cells.ShdwForegndTrans = row[this.ShdwForegndTrans];
            cells.ShdwPattern = Extensions.CellDataMethods.ToInt(row[this.ShdwPattern]);
            cells.BeginArrow = Extensions.CellDataMethods.ToInt(row[this.BeginArrow]);
            cells.BeginArrowSize = row[this.BeginArrowSize];
            cells.EndArrow = Extensions.CellDataMethods.ToInt(row[this.EndArrow]);
            cells.EndArrowSize = row[this.EndArrowSize];
            cells.LineCap = Extensions.CellDataMethods.ToInt(row[this.LineCap]);
            cells.LineColor = Extensions.CellDataMethods.ToInt(row[this.LineColor]);
            cells.LineColorTrans = row[this.LineColorTrans];
            cells.LinePattern = Extensions.CellDataMethods.ToInt(row[this.LinePattern]);
            cells.LineWeight = row[this.LineWeight];
            cells.Rounding = row[this.Rounding];
            return cells;
        }

    }
}