using SRCCON = VisioAutomation.ShapeSheet.SRCConstants;

namespace VisioAutomation.ShapeSheetQuery.CommonQueries
{
    class XFormCellsQuery : CellQuery
    {
        public ColumnSRC Width { get; set; }
        public ColumnSRC Height { get; set; }
        public ColumnSRC PinX { get; set; }
        public ColumnSRC PinY { get; set; }
        public ColumnSRC LocPinX { get; set; }
        public ColumnSRC LocPinY { get; set; }
        public ColumnSRC Angle { get; set; }

        public XFormCellsQuery()
        {
            this.PinX = this.AddCell(SRCCON.PinX, nameof(SRCCON.PinX));
            this.PinY = this.AddCell(SRCCON.PinY, nameof(SRCCON.PinY));
            this.LocPinX = this.AddCell(SRCCON.LocPinX, nameof(SRCCON.LocPinX));
            this.LocPinY = this.AddCell(SRCCON.LocPinY, nameof(SRCCON.LocPinY));
            this.Width = this.AddCell(SRCCON.Width, nameof(SRCCON.Width));
            this.Height = this.AddCell(SRCCON.Height, nameof(SRCCON.Height));
            this.Angle = this.AddCell(SRCCON.Angle, nameof(SRCCON.Angle));
        }

        public Shapes.XFormCells GetCells(ShapeSheet.CellData<double>[] row)
        {
            var cells = new Shapes.XFormCells
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