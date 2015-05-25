namespace VisioAutomation.ShapeSheet.Query.Common
{
    class XFormCellsQuery : CellQuery
    {
        public Query.CellColumn Width { get; set; }
        public Query.CellColumn Height { get; set; }
        public Query.CellColumn PinX { get; set; }
        public Query.CellColumn PinY { get; set; }
        public Query.CellColumn LocPinX { get; set; }
        public Query.CellColumn LocPinY { get; set; }
        public Query.CellColumn Angle { get; set; }

        public XFormCellsQuery()
        {
            this.PinX = this.AddCell(ShapeSheet.SRCConstants.PinX, nameof(ShapeSheet.SRCConstants.PinX));
            this.PinY = this.AddCell(ShapeSheet.SRCConstants.PinY, nameof(ShapeSheet.SRCConstants.PinY));
            this.LocPinX = this.AddCell(ShapeSheet.SRCConstants.LocPinX, nameof(ShapeSheet.SRCConstants.LocPinX));
            this.LocPinY = this.AddCell(ShapeSheet.SRCConstants.LocPinY, nameof(ShapeSheet.SRCConstants.LocPinY));
            this.Width = this.AddCell(ShapeSheet.SRCConstants.Width, nameof(ShapeSheet.SRCConstants.Width));
            this.Height = this.AddCell(ShapeSheet.SRCConstants.Height, nameof(ShapeSheet.SRCConstants.Height));
            this.Angle = this.AddCell(ShapeSheet.SRCConstants.Angle, nameof(ShapeSheet.SRCConstants.Angle));
        }

        public VisioAutomation.Shapes.XFormCells GetCells(System.Collections.Generic.IList<ShapeSheet.CellData<double>> row)
        {
            var cells = new VisioAutomation.Shapes.XFormCells
            {
                PinX = row[this.PinX],
                PinY = row[this.PinY],
                LocPinX = row[this.LocPinX],
                LocPinY = row[this.LocPinY],
                Width = row[this.Width],
                Height = row[this.Height],
                Angle = row[this.Angle]
            };
            return cells;
        }
    }
}