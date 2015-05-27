using SRCCON = VisioAutomation.ShapeSheet.SRCConstants;

namespace VisioAutomation.ShapeSheet.Query.Common
{
    class ShapeFormatCellsQuery : CellQuery
    {
        public Query.CellColumn FillBkgnd { get; set; }
        public Query.CellColumn FillBkgndTrans { get; set; }
        public Query.CellColumn FillForegnd { get; set; }
        public Query.CellColumn FillForegndTrans { get; set; }
        public Query.CellColumn FillPattern { get; set; }
        public Query.CellColumn ShapeShdwObliqueAngle { get; set; }
        public Query.CellColumn ShapeShdwOffsetX { get; set; }
        public Query.CellColumn ShapeShdwOffsetY { get; set; }
        public Query.CellColumn ShapeShdwScaleFactor { get; set; }
        public Query.CellColumn ShapeShdwType { get; set; }
        public Query.CellColumn ShdwBkgnd { get; set; }
        public Query.CellColumn ShdwBkgndTrans { get; set; }
        public Query.CellColumn ShdwForegnd { get; set; }
        public Query.CellColumn ShdwForegndTrans { get; set; }
        public Query.CellColumn ShdwPattern { get; set; }
        public Query.CellColumn BeginArrow { get; set; }
        public Query.CellColumn BeginArrowSize { get; set; }
        public Query.CellColumn EndArrow { get; set; }
        public Query.CellColumn EndArrowSize { get; set; }
        public Query.CellColumn LineColor { get; set; }
        public Query.CellColumn LineCap { get; set; }
        public Query.CellColumn LineColorTrans { get; set; }
        public Query.CellColumn LinePattern { get; set; }
        public Query.CellColumn LineWeight { get; set; }
        public Query.CellColumn Rounding { get; set; }

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

        public VisioAutomation.Shapes.FormatCells GetCells(System.Collections.Generic.IList<ShapeSheet.CellData<double>> row)
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