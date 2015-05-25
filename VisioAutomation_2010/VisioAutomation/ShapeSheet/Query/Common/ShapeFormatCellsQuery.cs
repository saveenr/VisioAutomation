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
            this.FillBkgnd = this.AddCell(ShapeSheet.SRCConstants.FillBkgnd, nameof(ShapeSheet.SRCConstants.FillBkgnd));
            this.FillBkgndTrans = this.AddCell(ShapeSheet.SRCConstants.FillBkgndTrans, nameof(ShapeSheet.SRCConstants.FillBkgndTrans));
            this.FillForegnd = this.AddCell(ShapeSheet.SRCConstants.FillForegnd, nameof(ShapeSheet.SRCConstants.FillForegnd));
            this.FillForegndTrans = this.AddCell(ShapeSheet.SRCConstants.FillForegndTrans, nameof(ShapeSheet.SRCConstants.FillForegndTrans));
            this.FillPattern = this.AddCell(ShapeSheet.SRCConstants.FillPattern, nameof(ShapeSheet.SRCConstants.FillPattern));
            this.ShapeShdwObliqueAngle = this.AddCell(ShapeSheet.SRCConstants.ShapeShdwObliqueAngle, nameof(ShapeSheet.SRCConstants.ShapeShdwObliqueAngle));
            this.ShapeShdwOffsetX = this.AddCell(ShapeSheet.SRCConstants.ShapeShdwOffsetX, nameof(ShapeSheet.SRCConstants.ShapeShdwOffsetX));
            this.ShapeShdwOffsetY = this.AddCell(ShapeSheet.SRCConstants.ShapeShdwOffsetY, nameof(ShapeSheet.SRCConstants.ShapeShdwOffsetY));
            this.ShapeShdwScaleFactor = this.AddCell(ShapeSheet.SRCConstants.ShapeShdwScaleFactor, nameof(ShapeSheet.SRCConstants.ShapeShdwScaleFactor));
            this.ShapeShdwType = this.AddCell(ShapeSheet.SRCConstants.ShapeShdwType, nameof(ShapeSheet.SRCConstants.ShapeShdwType));
            this.ShdwBkgnd = this.AddCell(ShapeSheet.SRCConstants.ShdwBkgnd, nameof(ShapeSheet.SRCConstants.ShdwBkgnd));
            this.ShdwBkgndTrans = this.AddCell(ShapeSheet.SRCConstants.ShdwBkgndTrans, nameof(ShapeSheet.SRCConstants.ShdwBkgndTrans));
            this.ShdwForegnd = this.AddCell(ShapeSheet.SRCConstants.ShdwForegnd, nameof(ShapeSheet.SRCConstants.ShdwForegnd));
            this.ShdwForegndTrans = this.AddCell(ShapeSheet.SRCConstants.ShdwForegndTrans, nameof(ShapeSheet.SRCConstants.ShdwForegndTrans));
            this.ShdwPattern = this.AddCell(ShapeSheet.SRCConstants.ShdwPattern, nameof(ShapeSheet.SRCConstants.ShdwPattern));
            this.BeginArrow = this.AddCell(ShapeSheet.SRCConstants.BeginArrow, nameof(ShapeSheet.SRCConstants.BeginArrow));
            this.BeginArrowSize = this.AddCell(ShapeSheet.SRCConstants.BeginArrowSize, nameof(ShapeSheet.SRCConstants.BeginArrowSize));
            this.EndArrow = this.AddCell(ShapeSheet.SRCConstants.EndArrow, nameof(ShapeSheet.SRCConstants.EndArrow));
            this.EndArrowSize = this.AddCell(ShapeSheet.SRCConstants.EndArrowSize, nameof(ShapeSheet.SRCConstants.EndArrowSize));
            this.LineColor = this.AddCell(ShapeSheet.SRCConstants.LineColor, nameof(ShapeSheet.SRCConstants.LineColor));
            this.LineCap = this.AddCell(ShapeSheet.SRCConstants.LineCap, nameof(ShapeSheet.SRCConstants.LineCap));
            this.LineColorTrans = this.AddCell(ShapeSheet.SRCConstants.LineColorTrans, nameof(ShapeSheet.SRCConstants.LineColorTrans));
            this.LinePattern = this.AddCell(ShapeSheet.SRCConstants.LinePattern, nameof(ShapeSheet.SRCConstants.LinePattern));
            this.LineWeight = this.AddCell(ShapeSheet.SRCConstants.LineWeight, nameof(ShapeSheet.SRCConstants.LineWeight));
            this.Rounding = this.AddCell(ShapeSheet.SRCConstants.Rounding, nameof(ShapeSheet.SRCConstants.Rounding));

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