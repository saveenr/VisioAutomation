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
            this.PinX = this.AddCell(ShapeSheet.SRCConstants.PinX, "PinX");
            this.PinY = this.AddCell(ShapeSheet.SRCConstants.PinY, "PinY");
            this.LocPinX = this.AddCell(ShapeSheet.SRCConstants.LocPinX, "LocPinX");
            this.LocPinY = this.AddCell(ShapeSheet.SRCConstants.LocPinY, "LocPinY");
            this.Width = this.AddCell(ShapeSheet.SRCConstants.Width, "Width");
            this.Height = this.AddCell(ShapeSheet.SRCConstants.Height, "Height");
            this.Angle = this.AddCell(ShapeSheet.SRCConstants.Angle, "Angle");
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